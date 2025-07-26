using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.medida;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion
{
       

    public partial class FrmXML_Respuestas_Datos_C105 : Form
    {
        public EndesaBusiness.cnmc.CNMC cnmc { get; set; }
        //  
        //   
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml { get; set; }

        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_A xml_a { get; set; }

        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_R xml_r { get; set; }

        //    public Dictionary<string, string> dic_tipo_activacion_prevista { get; set; }
        public EndesaBusiness.xml.ContratacionXML cont_xml { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC105 xml_c105 { get; set; }

        public List<string> lista_motivo_rechazo { get; set; }
        
        public DateTime fecha { get; set; }

        public FrmXML_Respuestas_Datos_C105()
        {
            InitializeComponent();
        }

       
        public FrmXML_Respuestas_Datos_C105(
                 EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml,
               EndesaBusiness.cnmc.CNMC cnmc,
              EndesaBusiness.xml.ContratacionXML cont_xml)
               : this()
        {
            this.xml = xml;
            this.cnmc = cnmc;
            this.cont_xml = cont_xml;
        }
        public class TipoContrato
        {
            public string Key { get; set; }         // Ej: "Anual"
            public string Display { get; set; }     // Ej: "01 - Anual"

            public override string ToString()
            {
                return $"{{ Key = {Key}, Display = {Display} }}";
            }
        }

        public class TipoValor
        {
            public string Key { get; set; }
            public string Display { get; set; }

            public override string ToString() => Display;
        }

        private void EvaluarVisibilidadFechaFinalizacion()
        {
            string tipoContratoATR = cmb_TipoContratoATR.Text?.Trim();

            if (string.IsNullOrWhiteSpace(tipoContratoATR) || tipoContratoATR.Length < 2)
            {
                // No hay valor válido seleccionado
                txt_fechafinalizacion.Visible = false;
                txt_fechafinalizacion.Enabled = false;
                txt_fechafinalizacion.Text = string.Empty;
                label8.Visible = false;
                return;
            }

            // Extraer los 2 últimos caracteres → ejemplo: "Anual - 01" → "01"
            tipoContratoATR = tipoContratoATR.Substring(tipoContratoATR.Length - 2);

           // MessageBox.Show($"TipoContratoATR = {tipoContratoATR}");

            bool mostrar = tipoContratoATR == "02" || tipoContratoATR == "03" || tipoContratoATR == "09";

            txt_fechafinalizacion.Visible = mostrar;
            txt_fechafinalizacion.Enabled = mostrar;
            label8.Visible = mostrar;

            if (!mostrar)
                txt_fechafinalizacion.Text = string.Empty;
        }




        private void cmb_TipoContratoATR_SelectedIndexChanged(object sender, EventArgs e)
        {
            EvaluarVisibilidadFechaFinalizacion();
        }
        private void EvaluarVisibilidadAutoconsumo()
        {
            string tipoAutoconsumo = cmb_Tipoautoconsumo.SelectedItem?.ToString()?.Substring(0, 2);
            groupBox4.Visible = tipoAutoconsumo != "00" && tipoAutoconsumo != "0C";
        }

        private void cmb_Tipoautoconsumo_SelectedIndexChanged(object sender, EventArgs e)
        {
            EvaluarVisibilidadAutoconsumo();
        }

        private void btnOK_Click(object sender, EventArgs e)
         {
            string displayText = cmb_TipoContratoATR.Text;
            string tipoContratoATR = displayText.Length >= 2
                ? displayText.Substring(displayText.Length - 2)
                : null;
            string fechaFinalizacion = null;

            // Validar solo si el campo está visible
            if (txt_fechafinalizacion.Visible)
            {
                if (string.IsNullOrWhiteSpace(txt_fechafinalizacion.Text))
                {
                    MessageBox.Show("Debe informar la Fecha de Finalización.",
                        "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!DateTime.TryParseExact(txt_fechafinalizacion.Text, "dd/MM/yyyy", null,
                    DateTimeStyles.None, out DateTime fechaParsed))
                {
                    MessageBox.Show("La Fecha de Finalización no tiene un formato válido (dd/MM/yyyy).",
                        "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                fechaFinalizacion = fechaParsed.ToString("yyyy-MM-dd");
            }
            //string tipoContratoATR = cmb_TipoContratoATR.SelectedValue?.ToString();

            string tipoAutoconsumo = cmb_Tipoautoconsumo.SelectedItem?.ToString()?.Substring(0, 2);
            

            bool hayAutoconsumoActivo = tipoAutoconsumo != "00" && tipoAutoconsumo != "0C";

            
            if (hayAutoconsumoActivo)
            {
                if (string.IsNullOrWhiteSpace(txt_cau.Text) ||
                    cmb_TipoCUPS.SelectedIndex == -1 ||
                    cmb_TipoSubseccion.SelectedIndex == -1 ||
                    cmb_TecGenerador.SelectedIndex == -1 ||
                    string.IsNullOrWhiteSpace(txt_PotInstaladaGen.Text) ||
                    cmb_TipoInstalacion.SelectedIndex == -1 ||
                    cmb_ssaa.SelectedIndex == -1)
                {
                    MessageBox.Show("Debe informar todos los campos de autoconsumo.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Crear objeto principal
            xml_c105 = new TipoMensajeC105();

            // Cabecera
            xml_c105.Cabecera = new Cabecera
            {
                CodigoREEEmpresaEmisora = xml.Cabecera.CodigoREEEmpresaDestino,
                CodigoREEEmpresaDestino = xml.Cabecera.CodigoREEEmpresaEmisora,
                CodigoDelProceso = xml.Cabecera.CodigoDelProceso,
                CodigoDePaso = "05",
                CodigoDeSolicitud = xml.Cabecera.CodigoDeSolicitud,
                SecuencialDeSolicitud = xml.Cabecera.SecuencialDeSolicitud,
                FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                CUPS = xml.Cabecera.CUPS
            };

            // Activación C105
            var activacion = new ActivacionCambiodeComercializadorSinCambios_C105();
            xml_c105.ActivacionCambiodeComercializadorSinCambios = activacion;
         

            // Datos Activación
            activacion.DatosActivacion = new DatosActivacionC105
            {
                Fecha = string.IsNullOrWhiteSpace(txt_fechaactivacion.Text)

                    ? null : DateTime.ParseExact(txt_fechaactivacion.Text, "dd/MM/yyyy", null).ToString("yyyy-MM-dd"),

                 IndEsencial = cmb_IndEsencial.SelectedItem?.ToString()?.Substring(0, 2),
                FechaUltimoMovimientoIndEsencial = "1900-01-01"
            };

            // Contrato
            var contrato = new Contrato_C105();
            contrato.IdContrato = new IdContrato
            {
                CodContrato = txt_codcontrato.Text
            };

           // contrato.FechaFinalizacion = fechaFinalizacion;

           //  contrato.FechaFinalizacion = string.IsNullOrWhiteSpace(txt_fechafinalizacion.Text)
            //    ? null : DateTime.ParseExact(txt_fechafinalizacion.Text, "dd/MM/yyyy", null).ToString("yyyy-MM-dd");

            contrato.TipoContratoATR = tipoContratoATR;

            if (!string.IsNullOrWhiteSpace(fechaFinalizacion))
            {
                contrato.FechaFinalizacion = fechaFinalizacion;
            }


            // Condiciones contractuales
            contrato.CondicionesContractuales = new CondicionesContractuales
            {
                TarifaATR = cmb_TarifaATR.SelectedValue?.ToString(),
                TensionDelSuministro = cmb_TensionDelSuministro.SelectedValue?.ToString(), 

                PeriodicidadFacturacion = cmb_PeriodicidadFacturacion.SelectedItem?.ToString()?.Substring(0, 2),
                TipoTelegestion = cmb_TipodeTelegestion.SelectedItem?.ToString()?.Substring(0, 2),

                PotenciasContratadas = new PotenciasContratadas()
            };

                        
            // Potencias contratadas
            foreach (DataGridViewRow row in dgvPotencias.Rows)
            {
                if (!row.IsNewRow)
                {
                    var periodo = row.Cells["Periodo"].Value?.ToString();
                    var valor = row.Cells["PotenciaW"].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(periodo) && !string.IsNullOrWhiteSpace(valor))
                    {
                        contrato.CondicionesContractuales.PotenciasContratadas.Potencia.Add(new Potencia
                        {
                            periodo = periodo,
                            potencia = valor
                        });
                    }
                }
            }

            // Autoconsumo si aplica
            if (hayAutoconsumoActivo)
            {
                contrato.Autoconsumo = new AutoconsumoSolicitudAlta
                {
                    DatosSuministro = new DatosSuministroSolicitud
                    {
                        TipoCUPS = cmb_TipoCUPS.SelectedItem?.ToString()?.Substring(0, 2)
                    },
                    DatosCAU = new DatosCAUAlta
                    {
                        CAU = txt_cau.Text,
                        TipoAutoconsumo = tipoAutoconsumo,
                        TipoSubseccion = cmb_TipoSubseccion.SelectedItem?.ToString()?.Substring(0, 2),
                     //   Colectivo = cmb_colectivo.SelectedItem?.ToString()?.Substring(0, 2),
                        DatosInstGen = new DatosInstGenSolicitud
                        {
                            TecGenerador = cmb_TecGenerador.SelectedItem?.ToString()?.Substring(0, 3),
                            PotInstaladaGen = txt_PotInstaladaGen.Text,
                            TipoInstalacion = cmb_TipoInstalacion.SelectedItem?.ToString()?.Substring(0, 2),
                            SSAA = cmb_ssaa.SelectedItem?.ToString()?.Substring(0, 1),
                            UnicoContrato = "1"
                        }
                    }
                };
            }

            activacion.Contrato = contrato;

            string codPM = ValidarCodPM(this.txt_CodPM); 
            if (codPM == null)
                return; // La validación falló, se mostró un mensaje y se aborta

            // Punto de medida
            activacion.PuntoDeMedida = new PuntoDeMedida
            {
                //CodPM = txt_CodPM.Text,
                CodPM = codPM,
                TipoMovimiento = cmb_TipoMovimiento.SelectedItem?.ToString()?.Substring(0, 1),
                TipoPM = cmb_tipoPM.SelectedItem?.ToString()?.Substring(0, 2),
                ModoLectura = cmd_Modolectura.SelectedItem?.ToString()?.Substring(0, 1),
                Funcion = cmb_Funcion.SelectedItem?.ToString()?.Substring(0, 1),
                // FechaVigor = string.IsNullOrWhiteSpace(Txt_FehaVigor.Text) ? null : DateTime.ParseExact(Txt_FehaVigor.Text, "dd/MM/yyyy", null).ToString("yyyy-MM-dd"),
                // FechaAlta = string.IsNullOrWhiteSpace(Txt_FechaAlta.Text) ? null : DateTime.ParseExact(Txt_FechaAlta.Text, "dd/MM/yyyy", null).ToString("yyyy-MM-dd")
                FechaVigor = ValidarYFormatearFecha(Txt_FehaVigor.Text, "Fecha de Vigor"),
                FechaAlta = ValidarYFormatearFecha(Txt_FechaAlta.Text, "Fecha de Alta"),
                //
                Aparatos= null // No se incluyen aparatos en C105
            };

            // Finalizar
            this.Tag = xml_c105;

            cont_xml.CreaMensajeC105(xml_c105);
          //  GuardarXML_C105(xml_c105);
            this.DialogResult = DialogResult.OK;
            this.Close();
            //;
            //
        }

        /// <summary>
        /// Valida que el contenido del TextBox contenga un CodPM válido (formato X(22)).
        /// Devuelve el código válido o null si no es válido.
        /// </summary>
        private string ValidarCodPM(Control control)
        {
            string codPM = control.Text?.Trim();

            if (string.IsNullOrWhiteSpace(codPM))
            {
                MessageBox.Show("Debe informar el Código de Punto de Medida (CodPM).", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                control.Focus();
                return null;
            }

            if (codPM.Length > 22)
            {
                MessageBox.Show("El Código de Punto de Medida no puede superar los 22 caracteres.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                control.Focus();
                return null;
            }

            if (!Regex.IsMatch(codPM, @"^[A-Za-z0-9]+$"))
            {
                MessageBox.Show("El Código de Punto de Medida solo puede contener letras y números.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                control.Focus();
                return null;
            }

            return codPM;
        }

        /// <summary>
        /// Valida y convierte una cadena de fecha en formato "dd/MM/yyyy" a "yyyy-MM-dd".
        /// Devuelve null si la cadena está vacía o si el formato es inválido.
        /// </summary>
        public static string ValidarYFormatearFecha(string fechaTexto, string nombreCampo = null)
        {
            if (string.IsNullOrWhiteSpace(fechaTexto))
                return null;

            fechaTexto = fechaTexto.Trim();

            if (DateTime.TryParseExact(fechaTexto, "dd/MM/yyyy", null,
                System.Globalization.DateTimeStyles.None, out DateTime fecha))
            {
                return fecha.ToString("yyyy-MM-dd");
            }
            else
            {
                if (!string.IsNullOrEmpty(nombreCampo))
                {
                    MessageBox.Show($"La fecha '{nombreCampo}' no tiene el formato correcto (dd/MM/yyyy).",
                                    "Formato de Fecha Incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return null;
            }
        }
        private void CargarCombosDesdeCNMC()
        {
            if (cnmc == null) return;

            // IndEsencial
            //cmb_IndEsencial.Items.Clear();
            //foreach (var kv in cnmc.dic_indicativo_sino)
            //    cmb_IndEsencial.Items.Add($"{kv.Key} - {kv.Value}");

            // TipoContratoATR
            cmb_TipoContratoATR.DataSource = cnmc.dic_contrato_atr
             .Select(kv => new TipoValor
             {
              Key = kv.Key,                       // lo que necesitas para el XML, como "01"
              Display = $"{kv.Key} - {kv.Value}"  // lo que el usuario ve: "01 - Anual"
              })
              .ToList();

            cmb_TipoContratoATR.DisplayMember = "Display";
            cmb_TipoContratoATR.ValueMember = "Key";
            cmb_TipoContratoATR.SelectedIndex = -1;


            // TarifaATR
            var tarifaIndexada = cnmc.dic_tarifa_atr
               .Select((kv, index) => new
               {
                Key = (index + 1).ToString("D2"),  // "01", "02", etc.
                Value = kv.Key,                    // "2.0TD", "3.0TD", etc.
                Display = $"{(index + 1).ToString("D2")} - {kv.Key}"
                })
               .ToList();

            cmb_TarifaATR.DataSource = tarifaIndexada;
            cmb_TarifaATR.DisplayMember = "Display";
            cmb_TarifaATR.ValueMember = "Key";

            // Tensión del suministro
            var tensionIndexada = cnmc.dic_tensiones
             .Select((kv, index) => new
             {
             Key = (index + 1).ToString("D2"),
             Value = kv.Key,
              Display = $"{(index + 1).ToString("D2")} - {kv.Key}"
             })
             .ToList();

            cmb_TensionDelSuministro.DataSource = tensionIndexada;
            cmb_TensionDelSuministro.DisplayMember = "Display";
            cmb_TensionDelSuministro.ValueMember = "Key";


            // Periodicidad facturación (manual)
            cmb_PeriodicidadFacturacion.Items.Clear();
            cmb_PeriodicidadFacturacion.Items.Add("01 - Mensual");
            cmb_PeriodicidadFacturacion.Items.Add("02 - Bimestral");
            //cmb_PeriodicidadFacturacion.Items.Add("03 - Trimestral");

            // Tipo de Telegestión (manual)
            cmb_TipodeTelegestion.Items.Clear();
            cmb_TipodeTelegestion.Items.Add("01 - Telegestión operativa con Curva de Carga Horaria");
            cmb_TipodeTelegestion.Items.Add("02 - Telegestión Operativa sin curva de carga Horaria");
            cmb_TipodeTelegestion.Items.Add("03 - Sin Telegestión");

            // Tipo Autoconsumo
            cmb_Tipoautoconsumo.Items.Clear();
            foreach (var kv in cnmc.dic_autoconsumo)
                cmb_Tipoautoconsumo.Items.Add($"{kv.Value} - {kv.Key}");
            //

            cmb_TipoContratoATR.SelectedIndex = -1;
            cmb_TarifaATR.SelectedIndex = -1;
            cmb_TensionDelSuministro.SelectedIndex = -1;

        }


        private void FrmXML_Respuestas_Fecha_Load(object sender, EventArgs e)
        {
            if (xml != null)
            {
                // Cabecera
                txtboxCUPS.Text = xml.Cabecera?.CUPS ?? "";
                txtboxCodigoSolicitud.Text = xml.Cabecera?.CodigoDeSolicitud ?? "";

            }

            // Ocultar bloques dependientes
             groupBox4.Visible = false;
             //label8.Visible = false;
             //txt_fechafinalizacion.Visible = false;  ;

            // Asignar eventos
            cmb_Tipoautoconsumo.SelectedIndexChanged += cmb_Tipoautoconsumo_SelectedIndexChanged;
            cmb_TipoContratoATR.SelectedIndexChanged += cmb_TipoContratoATR_SelectedIndexChanged;

            // Cargar datos
            CargarCombosDesdeCNMC();
            InicializarPotencias(); 

            
            // Opcional: seleccionar por defecto "00 - Sin autoconsumo"
            if (cmb_Tipoautoconsumo.Items.Count > 0)
            {
                var item = cmb_Tipoautoconsumo.Items
                    .Cast<string>()
                    .FirstOrDefault(x => x.StartsWith("00") || x.StartsWith("0C"));

                if (item != null) cmb_Tipoautoconsumo.SelectedItem = item;
            }

            if (cmb_IndEsencial.Items.Count > 0)
            {
                var item = cmb_IndEsencial.Items.Cast<string>().FirstOrDefault(x => x.StartsWith("00"));
                if (item != null) cmb_IndEsencial.SelectedItem = item;
            }

            //
            // MessageBox.Show($"TipoContratoATR: key = {cmb_TipoContratoATR.SelectedValue}, text = {cmb_TipoContratoATR.Text}");
            //{
            //    cmb_TipoContratoATR.SelectedValue = "02"; // o "03" o "09" según lo que necesites por defecto
            //}


            // Forzar visibilidad inicial de campos dependientes
            if (cmb_TipoContratoATR.SelectedIndex != -1)
            {
                EvaluarVisibilidadFechaFinalizacion();
            }

            EvaluarVisibilidadAutoconsumo();

        }
        private void InicializarPotencias()
        {
            dgvPotencias.Rows.Clear();

            for (int i = 1; i <= 6; i++)
            {
                dgvPotencias.Rows.Add($"0{i}", ""); // Periodo: 01 a 06
            }

            dgvPotencias.AllowUserToAddRows = false; // opcional, evitar fila vacía adicional
        }


        private void chk_rechazo_CheckedChanged(object sender, EventArgs e)
        {
          
        }

        private void chk_aceptacion_CheckedChanged(object sender, EventArgs e)
        {
            
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

        private void txt_codcontrato_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }
    }
}
