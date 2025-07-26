using EndesaBusiness.cnmc;
using EndesaEntity.cnmc;
using EndesaEntity.cnmc.V21_2019_12_17;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Contracts;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion
{
       

    public partial class FrmXML_Respuestas_Datos_B105 : Form
    {
        public EndesaBusiness.cnmc.CNMC cnmc { get; set; }
        // public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml { get; set; }
        public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml { get; set; }
        
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305 xml_a305 { get; set; }
        public EndesaBusiness.xml.ContratacionXML cont_xml { get; set; }
       // public Dictionary<string, string> dic_tipo_activacion_prevista { get; set; }
        //public List<string> lista_motivo_rechazo { get; set; }
        
        public DateTime fecha { get; set; }

        public FrmXML_Respuestas_Datos_B105()
        {
            InitializeComponent();

          //  chk_aceptacion.Checked = true;
         //   chk_actuacion_campo.Checked = true;
          
        }

        //private void btnOK_Click(object sender, EventArgs e)  //irh v1
        //{
        //    DateTime fecha;
        //    DateTime fechaUltimoMovimiento;

        //    if (!DateTime.TryParseExact(txt_fecha.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fecha) ||
        //        !DateTime.TryParseExact(txt_fechaultimomovimiento.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fechaUltimoMovimiento))
        //    {
        //        MessageBox.Show("Debe ingresar ambas fechas en formato dd/MM/yyyy", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return;
        //    }

        //    //}
        //    var xml_a305 = new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305();
        //    xml_a305.Cabecera = new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();

        //    // xml_a305.ActivacionAlta = new EndesaEntity.cnmc.V21_2019_12_17.ActivacionAlta();

        //    // Configurar cabecera
        //    xml_a305.Cabecera.CodigoREEEmpresaEmisora = xml.Cabecera.CodigoREEEmpresaDestino;
        //    xml_a305.Cabecera.CodigoREEEmpresaDestino = xml.Cabecera.CodigoREEEmpresaEmisora;
        //    xml_a305.Cabecera.CodigoDelProceso = xml.Cabecera.CodigoDelProceso;
        //    xml_a305.Cabecera.CodigoDePaso = "05";
        //    xml_a305.Cabecera.CodigoDeSolicitud = xml.Cabecera.CodigoDeSolicitud;
        //    xml_a305.Cabecera.SecuencialDeSolicitud = xml.Cabecera.SecuencialDeSolicitud;
        //    xml_a305.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
        //    xml_a305.Cabecera.CUPS = xml.Cabecera.CUPS;


        //    // ACTIVACION ALTA
        //    xml_a305.ActivacionAlta = new ActivacionAlta();
        //    xml_a305.ActivacionAlta.DatosActivacion = new DatosActivacion();
        //    xml_a305.ActivacionAlta.DatosActivacion.fecha = fecha.ToString("yyyy-MM-dd");
        //    xml_a305.ActivacionAlta.DatosActivacion.IndEsencial = cmb_indesencial.SelectedItem.ToString().Substring(0, 2);
        //    xml_a305.ActivacionAlta.DatosActivacion.FechaUltimoMovimientoIndEsencial = fechaUltimoMovimiento.ToString("yyyy-MM-dd");

        //    xml_a305.ActivacionAlta.Contrato = new ContratoAlta();
        //    xml_a305.ActivacionAlta.Contrato.IdContrato = new IdContrato();
        //    xml_a305.ActivacionAlta.Contrato.IdContrato.CodContrato = txt_codcontrato.Text;
        //    // Asignar TipoContratoATR
        //    xml_a305.ActivacionAlta.Contrato.TipoContratoATR = xml.Alta.Contrato.TipoContratoATR;

        //    string tc = xml.Alta.Contrato.TipoContratoATR;
        //    if (tc == "08" || tc == "11" || tc == "12")
        //    {
        //        string cupsPrincipal = xml.Cabecera.CUPS.Substring(0, 22);

        //        // Validamos que realmente son 22 caracteres
        //        if (cupsPrincipal.Length != 22)
        //        {
        //            throw new Exception("El CUPS extraído no tiene 22 caracteres: " + cupsPrincipal);
        //        }

        //        // Rellenamos con 3 ceros al final
        //        cupsPrincipal += "000";

        //        xml_a305.ActivacionAlta.Contrato.CUPSPrincipal = cupsPrincipal;
        //    }

        //    // Verificar si es autoconsumo activo

        //     if (xml.Alta.Contrato.Autoconsumo == null)
        //    {
        //        xml.Alta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();
        //    }
        //    string tipoAutoconsumo = xml.Alta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo;
        //    //var tipoautoconsumo = xml.Alta.Contrato.Autoconsumo?.DatosCAU?.TipoAutoconsumo;
        //    // ----AUTOCONSUMO ACTIVO----
        //    if (tipoAutoconsumo != "00" && tipoAutoconsumo != "0C")
        //    {
        //        //if (xml.Alta.Contrato.Autoconsumo == null)
        //        //{
        //        //xml.Alta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();
        //        //}

        //       if (xml.Alta.Contrato.Autoconsumo.DatosCAU == null)

        //        xml.Alta.Contrato.Autoconsumo.DatosCAU = new DatosCAUAlta();

        //        // DATOS SUMINISTRO

        //        InicializarAutoconsumo(xml_a305);

        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosSuministro.TipoCUPS = xml.Alta.Contrato.Autoconsumo.DatosSuministro.TipoCUPS;
        //        //  xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosSuministro.TipoCUPS = "01";


        //        // DATOS CAU---------------------------
        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.CAU = xml.Alta.Contrato.Autoconsumo.DatosCAU.CAU;
        //        //"ES0031609434759001ZP0FA666";//

        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo = xml.Alta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo;
        //        // xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo = "12";

        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.TipoSubseccion = xml.Alta.Contrato.Autoconsumo.DatosCAU.TipoSubseccion;
        //        // xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.TipoSubseccion = "10";

        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.Colectivo = xml.Alta.Contrato.Autoconsumo.DatosCAU.Colectivo;
        //        //xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.Colectivo = "S"; 

        //        // DATOS INSTGEN
        //        if (xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen == null)
        //        {
        //            xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen = new DatosInstGenSolicitud();
        //        }

        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen = new DatosInstGenSolicitud();
        //        // ojo no existe.  ver si se informa en form  es obligatorio 
        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TecGenerador = xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TecGenerador;
        //        //xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TecGenerador = "a11";

        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.PotInstaladaGen = xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.PotInstaladaGen;

        //        //xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.PotInstaladaGen = "5"; 

        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TipoInstalacion = xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TipoInstalacion;

        //        //xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TipoInstalacion = "01";


        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.SSAA = xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.SSAA;
        //        // xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.SSAA = "S";
        //        // xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.SSAA;
        //    }
        //    //----------  contacto  null
        //    xml_a305.ActivacionAlta.Contrato.Contacto = null;

        //    // --------------Inicializar CondicionesContractuales directamente
        //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales = new CondicionesContractuales();
        //    // Asignar TarifaATR
        //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TarifaATR =
        //        xml.Alta.Contrato.CondicionesContractuales.TarifaATR;
        //    // Asignar PeriodicidadFacturacion, valor por defecto "01" si está vacía  - - no existe en pas o01--- crear 
        //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PeriodicidadFacturacion = xml.Alta.Contrato.CondicionesContractuales.PeriodicidadFacturacion;

        //   // xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PeriodicidadFacturacion = "01";

        //    // Asignar TipoTelegestion (fijo)
        //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TipoTelegestion = xml.Alta.Contrato.CondicionesContractuales.TipoTelegestion;
        //    //xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TipoTelegestion = "01";

        //    // Asignar TensionDelSuministro  no esta en paso 01 02 -- crear
        //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TensionDelSuministro = xml.Alta.Contrato.CondicionesContractuales.TensionDelSuministro;
        //   // xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TensionDelSuministro = "01";


        //    // Inicializar y asignar PotenciasContratadas
        //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas = new PotenciasContratadas();
        //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia = new List<Potencia>();

        //    if (xml.Alta.Contrato.CondicionesContractuales.PotenciasContratadas?.Potencia != null)
        //    {
        //        foreach (var p in xml.Alta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia)
        //        {
        //            var potencia = new Potencia
        //            {
        //                periodo = p.periodo,
        //                potencia = p.potencia
        //            };

        //            xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia.Add(potencia);
        //        }
        //    }
        //    //-------------punto de medida    no existe        ----------------------
        //    //punto de medida - simulado 
        //    xml_a305.ActivacionAlta.PuntosDeMedida.PuntoDeMedida = new List<PuntoDeMedida>();

        //    PuntoDeMedida pm = new PuntoDeMedida(); // Crear una instancia de PuntoDeMedida
        //    pm.CodPM = "SVL1234567890123456789";
        //    pm.TipoMovimiento = "A";
        //    pm.TipoPM = "01";
        //    pm.ModoLectura = "1";
        //    pm.Funcion = "P";
        //    pm.FechaVigor = "2025-05-01";
        //    pm.FechaAlta = "2025-05-01";
        //    pm.Aparatos = null;

        //    xml_a305.ActivacionAlta.PuntosDeMedida.PuntoDeMedida.Add(pm);  // este objeto lo asigno al xml_a305 

        //    cont_xml.CreaMensajeA305v2(xml_a305);


        //    this.DialogResult = DialogResult.OK;
        //    this.Close();


        //}
        private void btnOK_Click(object sender, EventArgs e)
        {
            DateTime fecha;
            // DateTime fechaUltimoMovimiento;

            if (!DateTime.TryParseExact(txt_fecha.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fecha))
            //  !DateTime.TryParseExact(txt_fechaultimomovimiento.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fechaUltimoMovimiento))
            {
                MessageBox.Show("Debe ingresar ambas fechas en formato dd/MM/yyyy", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var xml_a305 = new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305();
            xml_a305.Cabecera = new EndesaEntity.cnmc.V21_2019_12_17.Cabecera();

            // Configurar cabecera
            xml_a305.Cabecera.CodigoREEEmpresaEmisora = xml.Cabecera.CodigoREEEmpresaDestino;
            xml_a305.Cabecera.CodigoREEEmpresaDestino = xml.Cabecera.CodigoREEEmpresaEmisora;
            xml_a305.Cabecera.CodigoDelProceso = xml.Cabecera.CodigoDelProceso;
            xml_a305.Cabecera.CodigoDePaso = "05";
            xml_a305.Cabecera.CodigoDeSolicitud = xml.Cabecera.CodigoDeSolicitud;
            xml_a305.Cabecera.SecuencialDeSolicitud = xml.Cabecera.SecuencialDeSolicitud;
            xml_a305.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            xml_a305.Cabecera.CUPS = xml.Cabecera.CUPS;

            // ACTIVACION ALTA
            xml_a305.ActivacionAlta = new ActivacionAlta();
            xml_a305.ActivacionAlta.DatosActivacion = new DatosActivacion(); 

            xml_a305.ActivacionAlta.DatosActivacion.fecha = fecha.ToString("yyyy-MM-dd");
            
            //  fehca finaja --
            xml_a305.ActivacionAlta.DatosActivacion.FechaUltimoMovimientoIndEsencial = "1999-01-01";
              
            //fechaUltimoMovimiento.ToString("yyyy-MM-dd");

            xml_a305.ActivacionAlta.Contrato = new ContratoAlta();
            xml_a305.ActivacionAlta.Contrato.IdContrato = new IdContrato();
            xml_a305.ActivacionAlta.Contrato.IdContrato.CodContrato = txt_codcontrato.Text;
            xml_a305.ActivacionAlta.Contrato.TipoContratoATR = xml.Alta.Contrato.TipoContratoATR;

            string tc = xml.Alta.Contrato.TipoContratoATR;
            if (tc == "08" || tc == "11" || tc == "12")
            {
                string cupsPrincipal = xml.Cabecera.CUPS.Substring(0, 22);
                if (cupsPrincipal.Length != 22)
                {
                    throw new Exception("El CUPS extraído no tiene 22 caracteres: " + cupsPrincipal);
                }

                cupsPrincipal += "000";
                xml_a305.ActivacionAlta.Contrato.CUPSPrincipal = cupsPrincipal;
            }

            // ---------------------------------------------------------
            // AUTOCONSUMO
            // ---------------------------------------------------------

            bool tieneAutoconsumo = xml.Alta?.Contrato?.Autoconsumo != null;
            string tipoAutoconsumo = xml.Alta?.Contrato?.Autoconsumo?.DatosCAU?.TipoAutoconsumo;

            bool esAutoconsumoActivo = !string.IsNullOrEmpty(tipoAutoconsumo)
                                        && tipoAutoconsumo != "00"
                                        && tipoAutoconsumo != "0C";

            if (esAutoconsumoActivo)
            {
                // Inicializar estructura Autoconsumo en xml_a305
                InicializarAutoconsumo(xml_a305);

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosSuministro.TipoCUPS =
                    xml.Alta.Contrato.Autoconsumo?.DatosSuministro?.TipoCUPS;

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.CAU =
                    xml.Alta.Contrato.Autoconsumo?.DatosCAU?.CAU;

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo =
                    xml.Alta.Contrato.Autoconsumo?.DatosCAU?.TipoAutoconsumo;

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.TipoSubseccion =
                    xml.Alta.Contrato.Autoconsumo?.DatosCAU?.TipoSubseccion;

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.Colectivo =
                    xml.Alta.Contrato.Autoconsumo?.DatosCAU?.Colectivo;
                //------------------------
                //DatosInstGen           - Obligatorio si TipoAutoconsumo de ese CAU <> "00" y  <> "0C"
                //--------------------------

                //xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TecGenerador = "a11";
                //  xml.Alta.Contrato.Autoconsumo?.DatosCAU?.DatosInstGen?.TecGenerador;
                //if (label17.Visible && cmb_TecGenerador.Visible)
                //{
                //    if (cmb_TecGenerador.SelectedItem != null)
                //    {
                //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TecGenerador =
                //            cmb_TecGenerador.SelectedItem.ToString().Substring(0, 2);
                //    }
                //    else
                //    {
                //        MessageBox.Show("Debe seleccionar un valor en Tec Generador.",
                //            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //        return;
                //    }
                //}
                //else
                //{
                //    xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TecGenerador = null;
                //}

                //
                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.PotInstaladaGen = 
                 xml.Alta.Contrato.Autoconsumo?.DatosCAU?.DatosInstGen?.PotInstaladaGen;

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.TipoInstalacion = 
                    xml.Alta.Contrato.Autoconsumo?.DatosCAU?.DatosInstGen?.TipoInstalacion;

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.SSAA =
                    xml.Alta.Contrato.Autoconsumo?.DatosCAU?.DatosInstGen?.SSAA;

                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen.UnicoContrato =
                    xml.Alta.Contrato.Autoconsumo?.DatosCAU?.DatosInstGen?.UnicoContrato;
            }
            else
            {
                // Si NO hay autoconsumo activo, no se inicializa el nodo en A305
                xml_a305.ActivacionAlta.Contrato.Autoconsumo = null;
            }

            // ---------- contacto null
            xml_a305.ActivacionAlta.Contrato.Contacto = null;

            // --------------CondicionesContractuales
            xml_a305.ActivacionAlta.Contrato.CondicionesContractuales = new CondicionesContractuales();
            xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TarifaATR =
                xml.Alta.Contrato.CondicionesContractuales?.TarifaATR;
            // form -  
           // xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PeriodicidadFacturacion = cmb_PeriodicidadFacturacion.SelectedItem.ToString().Substring(0, 2);
            //"01"; //vent
            // xml.Alta.Contrato.CondicionesContractuales?.PeriodicidadFacturacion;

            //xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TipoTelegestion = cmb_TipodeTelegestion.SelectedItem.ToString().Substring(0, 2);
            //"01"; //   vent  --ojo
            //  xml.Alta.Contrato.CondicionesContractuales?.TipoTelegestion;
            //-------------------------------------------------------------------------------------------------------
            //--- tabla  cnmc. dic_tensiones = Carga_Tabla_CNMC("cnmc_p_tensiones")

            //if (cmb_TensionDelSuministro.SelectedItem != null)
            //{
            //    var selected = (TensionSuministro)cmb_TensionDelSuministro.SelectedItem;

            //    string codigo = selected.Codigo;        // "01"
            //    string descripcion = selected.Descripcion; // "1X220"

            //    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TensionDelSuministro = descripcion;
            //    // Opcional: guardar también la descripción si quieres
            //    // xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.TensionDescripcion = descripcion;
            //}
            //else
            //{
            //    MessageBox.Show("Debe seleccionar la Tensión del Suministro.",
            //        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            //-------------------------------------------------

            // PotenciasContratadas
            xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas = new PotenciasContratadas();
            xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia = new List<Potencia>();

            if (xml.Alta.Contrato.CondicionesContractuales?.PotenciasContratadas?.Potencia != null)
            {
                foreach (var p in xml.Alta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia)
                {
                    var potencia = new Potencia
                    {
                        periodo = p.periodo,
                        potencia = p.potencia
                    };

                    xml_a305.ActivacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia.Add(potencia);
                }
            }


           // if (!DateTime.TryParseExact(txt_fecha.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fecha))


                // ------------ Inicialización de lista de PuntosDeMedida -------------------
                xml_a305.ActivacionAlta.PuntosDeMedida = new PuntosDeMedida();
            xml_a305.ActivacionAlta.PuntosDeMedida.PuntoDeMedida = new List<PuntoDeMedida>();

            // ------------ Validación de fechas -------------------

            DateTime fechaVigor;
            DateTime fechaAlta;

            if (string.IsNullOrWhiteSpace(Txt_FehaVigor.Text) ||
                !DateTime.TryParseExact(Txt_FehaVigor.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fechaVigor))
            {
                MessageBox.Show("Debe ingresar la Fecha Vigor en formato dd/MM/yyyy",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(Txt_FechaAlta.Text) ||
                !DateTime.TryParseExact(Txt_FechaAlta.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fechaAlta))
            {
                MessageBox.Show("Debe ingresar la Fecha Alta en formato dd/MM/yyyy",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------ Validación de CodPM -------------------

            string codPM = Txt_CodPM.Text?.Trim();

            if (string.IsNullOrEmpty(codPM))
            {
                MessageBox.Show("Debe ingresar el Código PM.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (codPM.Length != 22)
            {
                MessageBox.Show("El Código PM debe tener exactamente 22 caracteres.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!System.Text.RegularExpressions.Regex.IsMatch(codPM, @"^[A-Za-z0-9]{22}$"))
            {
                MessageBox.Show("El Código PM solo puede contener letras y números.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------ Creación del PuntoDeMedida -------------------

            PuntoDeMedida pm = new PuntoDeMedida
            {
                CodPM = codPM,
                TipoMovimiento = cmb_TipoMovimiento.SelectedItem.ToString().Substring(0, 1),
                TipoPM = cmb_tipoPM.SelectedItem.ToString().Substring(0, 2),
                ModoLectura = cmd_Modolectura.SelectedItem.ToString().Substring(0, 1),
                Funcion = cmb_Funcion.Text.Substring(0, 1),
                FechaVigor = fechaVigor.ToString("yyyy-MM-dd"),
                FechaAlta = fechaAlta.ToString("yyyy-MM-dd"),
                Aparatos = null
            };

            // ------------ Añadirlo a la lista -------------------

            xml_a305.ActivacionAlta.PuntosDeMedida.PuntoDeMedida.Add(pm);

            // Llamada al método para crear el mensaje A305

            cont_xml.CreaMensajeA305v2(xml_a305);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        private void InicializarAutoconsumo(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305 xml_a305)
        {
            if (xml_a305.ActivacionAlta == null)
                xml_a305.ActivacionAlta = new ActivacionAlta();

            if (xml_a305.ActivacionAlta.Contrato == null)
                xml_a305.ActivacionAlta.Contrato = new ContratoAlta();

            if (xml_a305.ActivacionAlta.Contrato.Autoconsumo == null)
                xml_a305.ActivacionAlta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();

            if (xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosSuministro == null)
                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosSuministro = new DatosSuministroSolicitud();

            if (xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU == null)
                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU = new DatosCAUAlta();

            if (xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen == null)
                xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen = new DatosInstGenSolicitud();
        }

        //private void InicializarAutoconsumo(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305 xml_a305)  //irh v
        //{
        //    if (xml_a305.ActivacionAlta == null)
        //        xml_a305.ActivacionAlta = new ActivacionAlta();

        //    if (xml_a305.ActivacionAlta.Contrato == null)
        //        xml_a305.ActivacionAlta.Contrato = new ContratoAlta();

        //    if (xml_a305.ActivacionAlta.Contrato.Autoconsumo == null)
        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();

        //    if (xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosSuministro == null)
        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosSuministro = new DatosSuministroSolicitud();

        //    if (xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU == null)
        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU = new DatosCAUAlta();

        //    if (xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen == null)
        //        xml_a305.ActivacionAlta.Contrato.Autoconsumo.DatosCAU.DatosInstGen = new DatosInstGenSolicitud();
        //}

        private void FrmXML_Respuestas_Datos_A305_Load(object sender, EventArgs e)
        {
            //irh
            txtboxCUPS.Text = xml.Cabecera.CUPS;
            txtboxCodigoSolicitud.Text = xml.Cabecera.CodigoDeSolicitud;
            // Cargar dinámicamente la Tensión del Suministro
         //   cmb_TensionDelSuministro.Items.Clear();

            //if (cnmc?.dic_tensiones != null && cnmc.dic_tensiones.Any())
            //{
            //    foreach (var kvp in cnmc.dic_tensiones)
            //    {
            //        cmb_TensionDelSuministro.Items.Add(
            //            new TensionSuministro(kvp.Key, kvp.Value));
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("No se encontraron datos de Tensión del Suministro en la base de datos.",
            //        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}


            // Verificar TipoAutoconsumo
            string tipoAutoconsumo = xml.Alta?.Contrato?.Autoconsumo?.DatosCAU?.TipoAutoconsumo;

            bool mostrarTecGenerador =
                !string.IsNullOrEmpty(tipoAutoconsumo) &&
                tipoAutoconsumo != "00" &&
                tipoAutoconsumo != "0C";

         //   label17.Visible = mostrarTecGenerador;
         //   cmb_TecGenerador.Visible = mostrarTecGenerador;
          
        }

        public class TensionSuministro
        {
            public string Codigo { get; set; }
            public string Descripcion { get; set; }

            public TensionSuministro(string codigo, string descripcion)
            {
                Codigo = codigo;
                Descripcion = descripcion;
            }

            public override string ToString()
            {
                return $"{Codigo} - {Descripcion}";
            }
        }

        private void FrmXML_Respuestas_Datos_A305_FormClosing(object sender, FormClosingEventArgs e)
        {
                        
            DialogResult result = MessageBox.Show("¿Deseas cerrar el formulario?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
                e.Cancel = true; // Cancela el cierre del formulario
            }
            else
            {
                this.DialogResult = DialogResult.Cancel;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txt_codcontrato_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            string input = txt_codcontrato.Text;

            // Comprobar si contiene solo números y tiene máximo 12 caracteres
            if (!System.Text.RegularExpressions.Regex.IsMatch(input, @"^\d{0,12}$"))
            {
                MessageBox.Show("El campo debe contener solo números y tener hasta 12 dígitos.");

                // Eliminar el último carácter ingresado no válido
                txt_codcontrato.Text = input.Remove(input.Length - 1);

                txt_codcontrato.SelectionStart = txt_codcontrato.Text.Length;
            }
        }

        private void gbDatosSolicitud_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }

     
}
