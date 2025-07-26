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
    public partial class FrmInformePaso_a_COR : Form
    {
        public FrmInformePaso_a_COR()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmInformePaso_a_COR_Load(object sender, EventArgs e)
        {
            DateTime fd;
            DateTime fh;

            fh = DateTime.Now.Date;
            fd = new DateTime(fh.Year, fh.Month, 1);

            txt_fd.Value = fd;
            txt_fh.Value = fh;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            EndesaBusiness.contratacion.eexxi.Aviso_a_COR cor =
                new EndesaBusiness.contratacion.eexxi.Aviso_a_COR();

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                cor.GeneraExcelRespuestaCOR(save.FileName, txt_fd.Value, txt_fh.Value);
                Cursor.Current = Cursors.Default;

                DialogResult result2 = MessageBox.Show("¿Desea abrir el Excel?", "Abrir Excel generado",
                   MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(save.FileName);
                }
            }


        }
    }
}
