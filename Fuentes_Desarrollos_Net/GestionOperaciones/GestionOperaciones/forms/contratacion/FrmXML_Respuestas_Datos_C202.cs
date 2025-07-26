using EndesaEntity.cnmc.V21_2019_12_17;
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
       

    public partial class FrmXML_Respuestas_Datos_C202 : Form
    {
        public EndesaBusiness.cnmc.CNMC cnmc { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml { get; set; }
  
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_A xml_a { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_R xml_r { get; set; }
       
        public EndesaBusiness.xml.ContratacionXML cont_xml { get; set; }
        public Dictionary<string, string> dic_tipo_activacion_prevista { get; set; }
        public List<string> lista_motivo_rechazo { get; set; }
        
        public DateTime fecha { get; set; }

        public FrmXML_Respuestas_Datos_C202()
        {
            InitializeComponent();
            this.AutoValidate = AutoValidate.Disable;
        }

         private void btnOK_Click(object sender, EventArgs e)
         {
            xml_a = new TipoMensajeC102_A();

            string codigoTarifa = cmb_TarifaATR.SelectedValue?.ToString() ?? "";
            MessageBox.Show($"Código tarifa seleccionado: {codigoTarifa}");

            var potenciasContratadas = new PotenciasContratadas();

            foreach (DataGridViewRow row in dgvPotencias.Rows)
            {
                if (!row.IsNewRow)
                {
                    string potenciaStr = row.Cells["PotenciaW"].Value?.ToString().Trim() ?? "";

                    if (!string.IsNullOrEmpty(potenciaStr))
                    {
                        if (decimal.TryParse(potenciaStr, out decimal potenciaDecimal))
                        {
                            if (potenciaDecimal < 0 || potenciaDecimal > 99999999999999)
                            {
                                MessageBox.Show("La potencia debe ser positiva y tener como máximo 14 dígitos.");
                                return;
                            }

                            var periodoStr = row.Cells["Periodo"].Value?.ToString() ?? "";

                            potenciasContratadas.Potencia.Add(new Potencia
                            {
                                periodo = periodoStr,
                                potencia = potenciaDecimal.ToString()
                            });
                        }
                        else
                        {
                            MessageBox.Show($"El valor '{potenciaStr}' no es numérico.");
                            return;
                        }
                    }
                }
            }

            if (potenciasContratadas.Potencia.Count > 6)
            {
                MessageBox.Show("Solo se permiten máximo 6 potencias contratadas.");
                return;
            }

            // Aquí continuarías con la serialización o guardado de datos...

            this.DialogResult = DialogResult.OK;
            this.Close();

            // Ejemplo: crear contrato y serializarlo
            //var contrato = new EndesaEntity.cnmc.V21_2019_12_17.Contrato
            //{
            //    TipoContratoATR = cmb_TipoContratoATR.SelectedValue?.ToString() ?? "",
            //    CondicionesContractuales = new EndesaEntity.cnmc.V21_2019_12_17.CondicionesContractuales
            //    {
            //        PotenciasContratadas = potenciasContratadas,
            //        TarifaATR = cmb_TarifaATR.SelectedValue?.ToString() ?? ""
            //    },
            //    Contacto = "..."
            //};

            //var serializer = new System.Xml.Serialization.XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.Contrato));
            //using (var writer = new System.IO.StreamWriter(@"C:\Temp\contrato.xml"))
            //{
            //    serializer.Serialize(writer, contrato);
            //}

        }

               

        private void FrmXML_Respuestas_Fecha_Load(object sender, EventArgs e)
        {


            txtboxCUPS.Text = xml.Cabecera.CUPS;
            txtboxCodigoSolicitud.Text = xml.Cabecera.CodigoDeSolicitud;

            //
            chk_aceptacion.ForeColor = Color.Green;
            chk_aceptacion.Font = new Font(chk_aceptacion.Font.FontFamily, 10, FontStyle.Bold);

            chk_rechazo.ForeColor = Color.Red;
            chk_rechazo.Font = new Font(chk_rechazo.Font.FontFamily, 10, FontStyle.Bold);

            if (cnmc?.dic_tarifa_atr != null)
            {
                var diccionarioConVacio1 = new Dictionary<string, string>
                {
                   { "", "" }
                };
                foreach (var kvp in cnmc.dic_tarifa_atr)
                {
                    diccionarioConVacio1.Add(kvp.Key, kvp.Value);
                }

                var lista = diccionarioConVacio1
                 .Select(kvp => new
                 {
                          Display = $"{kvp.Key} - {kvp.Value}",
                          Key = kvp.Key,
                           Value = kvp.Value
                 })
                  .ToList();

                 cmb_TarifaATR.DataSource = lista;
                 cmb_TarifaATR.DisplayMember = "Display"; // Esto se muestra en el combo
                 cmb_TarifaATR.ValueMember = "Value";     // Este es el valor seleccionado
            }
            if (cnmc?.dic_contrato_atr != null)
            {
                var diccionarioConVacio = new Dictionary<string, string>
                {
                   { "", "" }
                };

                foreach (var kvp in cnmc.dic_contrato_atr)
                {
                    diccionarioConVacio.Add(kvp.Key, kvp.Value);
                }

                cmb_TipoContratoATR.DataSource = new BindingSource(diccionarioConVacio, null);
                cmb_TipoContratoATR.DisplayMember = "Key";
                cmb_TipoContratoATR.ValueMember = "Value";
            }
            //----------------
            if (cnmc?.dic_tipo_activacion_prevista != null)
            {
                var diccionarioConVacio = new Dictionary<string, string>
                {
                 { "", "" }
                };

                foreach (var kvp in cnmc.dic_tipo_activacion_prevista)
                {
                    diccionarioConVacio.Add(kvp.Key, kvp.Value);
                }

                cmb_TipoActivacionPrevista.DataSource = new BindingSource(diccionarioConVacio, null);
                cmb_TipoActivacionPrevista.DisplayMember = "Key";
                cmb_TipoActivacionPrevista.ValueMember = "Value";
            }


            // Inicializar dgvPotencias con 6 periodos
            dgvPotencias.Rows.Clear();
            for (int i = 1; i <= 6; i++)
            {
                dgvPotencias.Rows.Add(i.ToString(), "");
            }


        }

        private void chk_rechazo_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_rechazo.Checked)
            {
                chk_aceptacion.Enabled = false;
                chk_aceptacion.Checked = false;
            }
            else
            {
                chk_aceptacion.Enabled = true;
            }
            if (chk_rechazo.Checked)
            {
                // Abrir la nueva ventana
                FrmXML_Respuestas_Datos_C102R nuevaVentana = new FrmXML_Respuestas_Datos_C102R();

                //  nuevaVentana.Show();
                nuevaVentana.ShowDialog(); //abrirla como ventana modal (bloqueante), usa .ShowDialog():
                // Cerrar la ventana actual
                this.Close();
            }

        }

        private void chk_aceptacion_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_aceptacion.Checked)
            {
                // Si se marca Aceptación → bloquear Rechazo
                chk_rechazo.Enabled = false;
                chk_rechazo.Checked = false; // Opcional: desmarcarlo si estuviera marcado
            }
            else
            {
                // Si se desmarca Aceptación → permitir Rechazo
                chk_rechazo.Enabled = true;
            }

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }
    }
}
