using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.facturacion
{
    public partial class FrmTipologiasPortugal : Form
    {

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmTipologiasPortugal()
        {

            usage.Start("Facturación", "FrmTipologiasPortugal" ,"N/A");
            InitializeComponent();
        }

        private void generarInformeToolStripMenuItem_Click(object sender, EventArgs e)
        {

            EndesaBusiness.facturacion.InformeInventarioTipologias_PT tipologiasPT =
                new EndesaBusiness.facturacion.InformeInventarioTipologias_PT();            
            
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            save.FileName = "Tipologias_PT_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                tipologiasPT.Inventario_por_Tipologias_v2(save.FileName);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }
            
        }

        private void opcionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Tipologías Portugal";
            p.tabla = "informe_inventario_tipologias_pt_param";
            p.esquemaString = "CON";
            p.Show();
        }

        private void FrmTipologiasPortugal_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Facturación", "FrmTipologiasPortugal" ,"N/A");
        }
    }
}
