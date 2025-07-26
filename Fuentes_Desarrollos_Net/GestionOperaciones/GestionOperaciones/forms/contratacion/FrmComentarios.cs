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

namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmComentarios : Form
    {

      
       
        EndesaBusiness.utilidades.Global g = new EndesaBusiness.utilidades.Global();
        EndesaBusiness.contratacion.eexxi.Comentarios comentarios;
        public string id_codigo { get; set; }
        
        public FrmComentarios()
        {


            InitializeComponent();                      

        }


        private void LoadData()
        {
            dgv.AutoGenerateColumns = false;
            comentarios = new EndesaBusiness.contratacion.eexxi.Comentarios(id_codigo);
            dgv.DataSource = comentarios.Lista_Comentarios(id_codigo);
        }
             

        private void FrmFacturasOperaciones_Load(object sender, EventArgs e)
        {
            LoadData();
        }
            

       
       
        private void cmdExcel_Click(object sender, EventArgs e)
        {
            DateTime ahora = new DateTime();
            
            int c = 0;

            Microsoft.Office.Interop.Excel.Application xlexcel;

            try
            {
                xlexcel = new Excel.Application();                

                ahora = DateTime.Now;
                Cursor = Cursors.WaitCursor;
               
                #region Informe Normal
                                        
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlexcel.ActiveSheet.Cells(1, 2).Value = "Complemento";                
                xlexcel.ActiveSheet.Cells(1, 3).Value = "Descripción";
                

                 c = 5;
                for (int i = 1; i <= c; i++)
                {
                    xlexcel.ActiveSheet.Cells(1, i).Font.Bold = true;
                }

                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[2, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                        
                 #endregion
                   
                 
            
            
            }catch(Exception ee)
            {
               MessageBox.Show(ee.Message,
               "Error en la construcción de la consulta",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
            Cursor = Cursors.Default;
            
        }

        
               
       
        private void btnEdit_Click(object sender, EventArgs e)
        {

            forms.contratacion.FrmComentariosEdit f = new FrmComentariosEdit();
            f.comentario = new EndesaEntity.contratacion.xxi.Comentario();
            f.editStatus = EndesaBusiness.contratacion.eexxi.Comentarios.EditStatus.EDIT;
            DataGridViewRow row = new DataGridViewRow();
            int fila = dgv.CurrentRow.Index;
            row = dgv.Rows[fila];
            f.comentario.id_comentario = row.Cells[0].Value.ToString();
            f.comentario.linea = Convert.ToInt32(row.Cells[1].Value);
            f.comentario.comentario = row.Cells[2].Value.ToString();
            f.txtComentario.Text = row.Cells[2].Value.ToString();
            f.comentario.created_by = row.Cells[3].Value.ToString();
            f.ShowDialog();
            LoadData();

        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            
            DialogResult r;

            try
            {
                DataGridViewRow row = new DataGridViewRow();
                int fila = dgv.CurrentRow.Index;
                row = dgv.Rows[fila];                

                r = MessageBox.Show("¿Desea borrar el comentario?",
                    "Borrar el comentario " + row.Cells[2].Value.ToString(),
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (r == DialogResult.Yes)
                {
                    comentarios.Del(row.Cells[0].Value.ToString(),
                        Convert.ToInt32(row.Cells[1].Value.ToString()));
                    LoadData();
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
        
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmComentariosEdit f = new FrmComentariosEdit();
            f.comentario = new EndesaEntity.contratacion.xxi.Comentario();
            f.editStatus = EndesaBusiness.contratacion.eexxi.Comentarios.EditStatus.NEW;            
            f.comentario.id_comentario = id_codigo;                        
            f.ShowDialog();
            LoadData();





        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
