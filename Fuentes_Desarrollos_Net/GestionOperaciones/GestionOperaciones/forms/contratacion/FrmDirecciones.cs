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
    public partial class FrmDirecciones : Form
    {
        public EndesaBusiness.contratacion.eexxi.EEXXI xxi { get; set; }
        public Dictionary<string, string> dic { get; set; }
        EndesaEntity.contratacion.xxi.XML_Datos xml;
        List<EndesaEntity.contratacion.xxi.XML_Datos> lista = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
        public FrmDirecciones()
        {
            InitializeComponent();
        }

        private void FrmDirecciones_Load(object sender, EventArgs e)
        {
            LoadData();
        }

       

        private void LoadData()
        {
            foreach (KeyValuePair<string, string> p in dic)
            {
                xml = xxi.BuscaDatosSolicitudXML("eexxi_solicitudes_tmp", p.Key, p.Value);
                if (!xml.encontrado_registro)
                    xml = xxi.BuscaDatosSolicitudXML("eexxi_solicitudes", p.Key, p.Value);

                lista.Add(xml);

            }

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = lista;
            //cboBoxColumn.DataSource = this.estados_casos.dic.Values.ToList();
        }


        private void BtnSave_Click(object sender, EventArgs e)
        {
            
            DialogResult result_1 = MessageBox.Show("¿Desea guardar los cambios?",
               "Cambios en direcciones", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result_1 == DialogResult.Yes)            
                Salvar();
            
               
        }

        private void Salvar()
        {
            EndesaBusiness.contratacion.eexxi.Solicitudes sol = new EndesaBusiness.contratacion.eexxi.Solicitudes();

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Cells[1].Value != null)
                {
                    sol.codigoDeSolicitud = row.Cells[0].Value.ToString();
                    sol.cups = row.Cells[4].Value.ToString();
                    if (row.Cells[5].Value != null)
                        sol.linea1DeLaDireccionExterna = row.Cells[5].Value.ToString();
                    else
                        sol.linea1DeLaDireccionExterna = null;
                    if (row.Cells[6].Value != null)
                        sol.linea2DeLaDireccionExterna = row.Cells[6].Value.ToString();
                    else
                        sol.linea2DeLaDireccionExterna = null;
                    if (row.Cells[7].Value != null)
                        sol.linea3DeLaDireccionExterna = row.Cells[7].Value.ToString();
                    else
                        sol.linea3DeLaDireccionExterna = null;
                    if (row.Cells[8].Value != null)
                        sol.linea4DeLaDireccionExterna = row.Cells[8].Value.ToString();
                    else
                        sol.linea4DeLaDireccionExterna = null;
                    sol.Update();
                }
            }
            LoadData();
            btnSave.Enabled = false;
        }

        private void Dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            this.dgv.Rows[e.RowIndex].Cells[1].Value = System.Environment.UserName.ToUpper();
        }

        private void Dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            btnSave.Enabled = true;
        }

        private void Dgv_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            // N/A
        }

        private void Dgv_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            btnSave.Enabled = true;
        }

        private void FrmDirecciones_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (btnSave.Enabled)
            {
                DialogResult result_1 = MessageBox.Show("¿Desea guardar los cambios?",
              "Cambios en direcciones", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result_1 == DialogResult.Yes)
                    Salvar();

            }
        }
               
    }
}
