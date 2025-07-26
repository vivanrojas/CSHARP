using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GestionOperaciones.forms
{
    
    public partial class FrmParameters : Form 
    {

        public string tabla;
        public string esquemaString;
        public string cabecera;
        


        DataTable dt;

        EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();
        EndesaBusiness.utilidades.Param p;

        
        public FrmParameters()
        {            
            InitializeComponent();
        }


        private void Cargadgv()
        {
           
            try
            {
                dgv.AutoGenerateColumns = false;
                dgv.DataSource = p.lista_parametros;                

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la carga de Parámetros",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }
        
        private void FrmParameters_Load(object sender, EventArgs e)
        {

            if (cabecera != null)
                this.Text = cabecera;

            switch (esquemaString)
            {
                case "FAC":
                    p = new EndesaBusiness.utilidades.Param(tabla, EndesaBusiness.servidores.MySQLDB.Esquemas.FAC);
                    break;
                case "MED":
                    p = new EndesaBusiness.utilidades.Param(tabla, EndesaBusiness.servidores.MySQLDB.Esquemas.MED);
                    break;
                case "CON":
                    p = new EndesaBusiness.utilidades.Param(tabla, EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                    break;
                case "AUX":
                    p = new EndesaBusiness.utilidades.Param(tabla, EndesaBusiness.servidores.MySQLDB.Esquemas.AUX);
                    break;
            }

            Cargadgv();
        }
               
        private void cmdExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Ubicación del informe";
            save.AddExtension = true;
            save.DefaultExt = "xlsx";
            save.Filter = "Ficheros xslx (*.xlsx)|*.*";
            DialogResult result = save.ShowDialog();
            if (result == DialogResult.OK)
            {
                p.ExportExcel(save.FileName);
                System.Diagnostics.Process.Start(save.FileName);
            }

        }

       

        private void listBoxEmpresas_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBoxLinea_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void txtCod_TextChanged(object sender, EventArgs e)
        {

        }
        
       

        private void btnAdd_Click(object sender, EventArgs e)
        {
                        
            forms.FrmParamEdit f = new forms.FrmParamEdit();
            f.Text = "Añadir nuevo registro";
            f.ShowDialog(this);

            if (f.txtCode.Text != null &&                 
                f.txtValue.Text != null)
            {
                p.code = f.txtCode.Text;
                p.from_date = Convert.ToDateTime(f.txt_begin_date.Text);
                p.to_date = Convert.ToDateTime(f.txt_to_date.Text);
                p.value = f.txtValue.Text;
                p.description = f.txtDescription.Text;

                p.Save();
                Cargadgv();
            }

           


        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            

            try
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgv.CurrentRow.Index;

                forms.FrmParamEdit f = new forms.FrmParamEdit();
                f.Text = "Editar registro";

                row = dgv.Rows[fila];
                f.txtCode.Text = row.Cells[0].Value.ToString();
                f.txt_begin_date.Text = row.Cells[1].Value.ToString();     
                f.txt_to_date.Text = row.Cells[2].Value.ToString();
                f.txtValue.Text = row.Cells[3].Value.ToString();
                f.txtDescription.Text = row.Cells[4].Value.ToString();

                f.ShowDialog(this);

                p.code = f.txtCode.Text;
                p.from_date = Convert.ToDateTime(f.txt_begin_date.Text);
                p.to_date = Convert.ToDateTime(f.txt_to_date.Text);
                p.value = f.txtValue.Text;
                p.description = f.txtDescription.Text;

                p.Save();
                Cargadgv();

                //if (f.txtCodigo.Text != cnae.complemento ||                    
                //    f.txtDescripcion.Text != cnae.descripcion)

                //{
                //    cnae.complemento = f.txtCodigo.Text;                    
                //    cnae.descripcion = f.txtDescripcion.Text;

                //    cnae.Update();
                //    Cargadgv();

                //}

            }
            catch(Exception ee)
            {
                MessageBox.Show(ee.Message,
               "Error en la edición de los datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
            
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            
            DialogResult r;

            try
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgv.CurrentRow.Index;
                row = dgv.Rows[fila];  

                r = MessageBox.Show("¿Desea borrar el registros con código " + row.Cells[0].Value.ToString() + "?",
                    "Borrar registro con código " + row.Cells[0].Value.ToString(),
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

               if (r == DialogResult.Yes)
                {
                    p.Delete(row.Cells[0].Value.ToString(), Convert.ToDateTime(row.Cells[1].Value), Convert.ToDateTime(row.Cells[2].Value));
                    Cargadgv();
                }
               
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message,
               "Error en el borrado de los datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
