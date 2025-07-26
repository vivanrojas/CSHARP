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
    public partial class FrmEndesaEnergia21_bajas_sin_altas : Form
    {
        EndesaBusiness.contratacion.eexxi.Solicitudes sol_b105;
        EndesaBusiness.contratacion.eexxi.Solicitudes sol_c106;
        EndesaBusiness.contratacion.eexxi.Solicitudes sol_c206;
        EndesaBusiness.contratacion.eexxi.Solicitudes sol_t101;
        EndesaBusiness.contratacion.eexxi.Solicitudes sol_t105;

        public FrmEndesaEnergia21_bajas_sin_altas()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmEndesaEnergia21_bajas_sin_altas_Load(object sender, EventArgs e)
        {
            EndesaEntity.contratacion.xxi.Cups_Solicitud o;

            List<EndesaEntity.contratacion.xxi.Cups_Solicitud> lista =
                new List<EndesaEntity.contratacion.xxi.Cups_Solicitud>();

            Cursor.Current = Cursors.WaitCursor;

            sol_b105 = new EndesaBusiness.contratacion.eexxi.Solicitudes("B1","05");
            sol_c106 = new EndesaBusiness.contratacion.eexxi.Solicitudes("C1", "06");
            sol_c206 = new EndesaBusiness.contratacion.eexxi.Solicitudes("C2", "06");
            sol_t101 = new EndesaBusiness.contratacion.eexxi.Solicitudes("T1", "01");
            sol_t105 = new EndesaBusiness.contratacion.eexxi.Solicitudes("T1", "05");

            foreach(KeyValuePair<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> p in sol_b105.dic)
            {
                if (!sol_t101.dic.TryGetValue(p.Key, out o))
                    lista.Add(p.Value);
                else if(!sol_t105.dic.TryGetValue(p.Key, out o))
                    lista.Add(p.Value);

            }

            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> p in sol_c106.dic)
            {
                if (!sol_t101.dic.TryGetValue(p.Key, out o))
                    lista.Add(p.Value);
                else if (!sol_t105.dic.TryGetValue(p.Key, out o))
                    lista.Add(p.Value);

            }

            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> p in sol_c206.dic)
            {
                if (!sol_t101.dic.TryGetValue(p.Key, out o))
                    lista.Add(p.Value);
                else if (!sol_t105.dic.TryGetValue(p.Key, out o))
                    lista.Add(p.Value);

            }

            lbl_total_solicitudes.Text = string.Format("Total Solicitudes: {0:#,##0}", lista.Count);

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;

            Cursor.Current = Cursors.Default;
        }
    }
}
