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
       

    public partial class FrmXML_Respuestas_Datos_C102 : Form
    {
        public EndesaBusiness.cnmc.CNMC cnmc { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml { get; set; }
  
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_A xml_a { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_R xml_r { get; set; }
       
        public EndesaBusiness.xml.ContratacionXML cont_xml { get; set; }
        public Dictionary<string, string> dic_tipo_activacion_prevista { get; set; }
        public List<string> lista_motivo_rechazo { get; set; }
        
        public DateTime fecha { get; set; }

        public FrmXML_Respuestas_Datos_C102()
        {
            InitializeComponent();
            this.AutoValidate = AutoValidate.Disable;
        }

        public FrmXML_Respuestas_Datos_C102(
                 EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml,
                 EndesaBusiness.cnmc.CNMC cnmc,
                 EndesaBusiness.xml.ContratacionXML cont_xml)
                  : this()
        {
            this.xml = xml;
            this.cnmc = cnmc;
            this.cont_xml = cont_xml;
        }
        private void btnOK_Click(object sender, EventArgs e)
         {
            xml_a = new TipoMensajeC102_A();
            // CABECERA → Se construye intercambiando emisora y destino
            xml_a.Cabecera.CodigoREEEmpresaEmisora = xml.Cabecera.CodigoREEEmpresaDestino;
            xml_a.Cabecera.CodigoREEEmpresaDestino = xml.Cabecera.CodigoREEEmpresaEmisora;
            xml_a.Cabecera.CodigoDelProceso = xml.Cabecera.CodigoDelProceso;
            xml_a.Cabecera.CodigoDePaso = "02";
            xml_a.Cabecera.CodigoDeSolicitud = xml.Cabecera.CodigoDeSolicitud;
            xml_a.Cabecera.SecuencialDeSolicitud = xml.Cabecera.SecuencialDeSolicitud;
            xml_a.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            xml_a.Cabecera.CUPS = xml.Cabecera.CUPS;

            // // DATOS ACEPTACION-----pte---
            if (!string.IsNullOrEmpty(txt_fecha_Aceptacion.Text))
            {
                xml_a.AceptacionCambiodeComercializadorSinCambios.DatosAceptacion.fechaAceptacion =

                  //  AceptacionAlta.DatosAceptacion.fechaAceptacion =
                    Convert.ToDateTime(txt_fecha_Aceptacion.Text).ToString("yyyy-MM-dd");
            }
            else
            {
                xml_a.AceptacionCambiodeComercializadorSinCambios.DatosAceptacion.fechaAceptacion = null;
                    
                    //.AceptacionAlta.DatosAceptacion.fechaAceptacion = null;
            }

            xml_a.AceptacionCambiodeComercializadorSinCambios.DatosAceptacion.ActuacionCampo = 
                //.AceptacionCambiodeComercializadorSinCambios.AceptacionAlta.DatosAceptacion.ActuacionCampo =
                cmb_ActuacionCampo.SelectedItem?.ToString().StartsWith("S") == true ? "S" : "N";
            // CONTRATO
            xml_a.AceptacionCambiodeComercializadorSinCambios.Contrato.TipoContratoATR =
                             
              //  .AceptacionAlta.Contrato.TipoContratoATR =
                cmb_TipoContratoATR.SelectedValue?.ToString();

            xml_a.AceptacionCambiodeComercializadorSinCambios.Contrato.TipoActivacionPrevista =
                
                
              //            .AceptacionAlta.Contrato.TipoActivacionPrevista =
                cmb_TipoActivacionPrevista.SelectedValue?.ToString();

            if (!string.IsNullOrEmpty(txt_FechaUltimaLecturaFirme.Text))
            {
                xml_a.AceptacionCambiodeComercializadorSinCambios.DatosAceptacion.FechaUltimaLecturaFirme =
                                    //                   .AceptacionAlta.DatosAceptacion.FechaUltimaLecturaFirme =
                      Convert.ToDateTime(txt_FechaUltimaLecturaFirme.Text).ToString("yyyy-MM-dd");
            }
            else
            {
                xml_a.AceptacionCambiodeComercializadorSinCambios.DatosAceptacion.FechaUltimaLecturaFirme = null;


                //  .AceptacionAlta.DatosAceptacion.FechaUltimaLecturaFirme = null;
            }

            if (!string.IsNullOrEmpty(txt_FechaActivacionPrevista.Text))
            {
                xml_a.AceptacionCambiodeComercializadorSinCambios.Contrato.FechaActivacionPrevista =

                    //                    .AceptacionAlta.Contrato.FechaActivacionPrevista =
                    Convert.ToDateTime(txt_FechaActivacionPrevista.Text).ToString("yyyy-MM-dd");
            }
            else
            {
                xml_a.AceptacionCambiodeComercializadorSinCambios.Contrato.FechaActivacionPrevista = null;


                //    .AceptacionAlta.Contrato.FechaActivacionPrevista = null;
            }

            xml_a.AceptacionCambiodeComercializadorSinCambios.Contrato.CondicionesContractuales.TarifaATR =


               // .AceptacionAlta.Contrato.CondicionesContractuales.TarifaATR =
                cmb_TarifaATR.SelectedValue?.ToString();


            // POTENCIAS
            var potenciasContratadas = new PotenciasContratadas();

            foreach (DataGridViewRow row in dgvPotencias.Rows)
            {
                if (!row.IsNewRow)
                {
                    string potenciaStr = row.Cells["PotenciaW"].Value?.ToString().Trim() ?? "";
                    string periodoStr = row.Cells["Periodo"].Value?.ToString();

                    if (!string.IsNullOrEmpty(potenciaStr))
                    {
                        if (decimal.TryParse(potenciaStr, out decimal potenciaDecimal))
                        {
                            if (potenciaDecimal < 0 || potenciaDecimal > 99999999999999)
                            {
                                MessageBox.Show("La potencia debe ser positiva y tener como máximo 14 dígitos.");
                                return;
                            }

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

            xml_a.AceptacionCambiodeComercializadorSinCambios.Contrato.CondicionesContractuales.PotenciasContratadas =
                   
                // xml_a.AceptacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas =
                potenciasContratadas;

            // TODO: Aquí llamarías a tu método para validar XSD y generar XML:
             cont_xml.CreaMensajeC102_A(xml_a);

            this.DialogResult = DialogResult.OK;
            this.Close();


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

                // Cargar lista de motivos desde CNMC
                this.lista_motivo_rechazo = cnmc.GetLista_Motivos_Rechazo("C1");

                // Abrir la nueva ventana
                FrmXML_Respuestas_Datos_C102R nuevaVentana =
                        new FrmXML_Respuestas_Datos_C102R(this.xml, this.cnmc, this.cont_xml, this.lista_motivo_rechazo);

                //FrmXML_Respuestas_Datos_C102R nuevaVentana = new FrmXML_Respuestas_Datos_C102R();

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
