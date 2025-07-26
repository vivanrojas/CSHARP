using EndesaBusiness.sharePoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones
{
    public partial class FormStart : Form
    {
        public string version = "04.06.2025";

        private Dictionary<string, string> lista_parametros =
            new Dictionary<string, string>();

        EndesaBusiness.logs.Log ficheroLog = new EndesaBusiness.logs.Log(Environment.CurrentDirectory, "logs", "Ini");
        
        public FormStart()
        {
            InitializeComponent();           
                        
        }
        private void facturasOperacionesToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            forms.facturacion.FrmFacturasOperaciones f = new forms.facturacion.FrmFacturasOperaciones();
            f.Show();            
        }

        private void FormStart_Load(object sender, EventArgs e)
        {
            ficheroLog.Add(" inicia sesion.");
            CargaParametros();
            toolStripStatusLabel1.Text = System.Environment.UserName +  " Versión: " + version;

            string server = GetValue("server_mysql_siope");

            this.Text = (server == "rdssiopepro.endesa.es" ?
                "GO Gestión de Operaciones (PRODUCCIÓN) " + server
                : "GO Gestión de Operaciones (DESARROLLO) " + server);

        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void facturasADIFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmAdifFacturas f =
                new forms.facturacion.FrmAdifFacturas();
            f.Show();
        }

        private void tamComparaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmTAM t = new forms.FrmTAM();
            //t.Show();
        }

        private void historicoFacturaciónReducidoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmFacturasReducido f = new forms.FrmFacturasReducido();
            //f.Show();
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {

        }

        private void cNAEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "c70ccnae_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros CNAE";
            p.Show();
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmDisuasoria f = new forms.contratacion.FrmDisuasoria();
            f.Show();
        }

        private void festivosEléctricosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmFacFestivosElectricosParam f = new forms.FrmFacFestivosElectricosParam();
            //f.tabla = "fact.fact_diasfestivos";
            //f.Text = "Festivos Eléctricos España";
            //f.Show();
        }

        private void cuadroDeMandoAGORAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmCuadroMandoAgora f = new forms.FrmCuadroMandoAgora();
            //f.Show();
        }

        private void acercaDeGOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.AboutForm f = new forms.AboutForm();
            f.ShowDialog();
        }

        private void gestiónDeCurvasDeCargaSCEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmCurvas f = new forms.medida.FrmCurvas();
            f.Show();
        }

        private void facturasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmSIGAME_Importe_Redes f = new forms.facturacion.FrmSIGAME_Importe_Redes();
            f.Show();
        }

        private void toolStripTextBox1_Click_1(object sender, EventArgs e)
        {

        }

        private void puntosSofisticadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmPuntosSofisticados f
                = new forms.facturacion.FrmPuntosSofisticados();
            f.Show();
        }

        private void calculadoraGasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmCalculadoraGas f = new forms.FrmCalculadoraGas();
            //f.Show();
        }

        private void impuestosToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void pNTsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmMedidaPNTs f = new forms.FrmMedidaPNTs();
            //f.Show();
        }

        private void pSATToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmPSAT f = new forms.FrmPSAT();
            //f.Show();
            forms.contratacion.FrmPS_AT_HIST f = new forms.contratacion.FrmPS_AT_HIST();
            f.Show();
        }

        private void informesERTEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmInformeERSE f = new forms.facturacion.FrmInformeERSE();
            f.Show();
        }

        private void aDIFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.medida.FrmAdif f = new forms.medida.FrmAdif();
            //f.Show();
        }

        private void cargaDeBloquesDeEnergíaParaFacturaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //OpenFileDialog d = new OpenFileDialog();
            //d.Filter = "txt files|*.txt";
            //d.Multiselect = false;
            //facturacion.bloques.BloquesFunciones bf = new facturacion.bloques.BloquesFunciones();
            //if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    foreach (string fileName in d.FileNames)
            //    {
            //        bf.CargaBloques(fileName);
            //    }

            //    if (bf.cargaOK)
            //    {
            //        MessageBox.Show("La importación ha concluido correctamente.",
            //       "Importación ficheros bloques de energía para facturas",
            //       MessageBoxButtons.OK,
            //       MessageBoxIcon.Information);
            //    }
                             
                
            //}
        }

        private void FormStart_FormClosing(object sender, FormClosingEventArgs e)
        {
            EndesaBusiness.utilidades.Global util = new EndesaBusiness.utilidades.Global();
            util.CerrarVentana();

        }

        private void distribuidorasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.gas.FrmDistribuidoras f = new forms.contratacion.gas.FrmDistribuidoras();
            f.Show();
        }

        private void importaciónPO1011ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.medida.Frm_PO1011 f = new forms.medida.Frm_PO1011();
            //f.Show();
        }

        private void estadosCargasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.herramientas.FrmCargas f = new forms.herramientas.FrmCargas();
            f.Show();
        }

        private void peajesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.medida.FrmPeajes f = new forms.medida.FrmPeajes();
            //f.Show();
        }

        private void importaciónCurvasDataMartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //GO.forms.medida.FrmImportaCurvasDataMart f = new forms.medida.FrmImportaCurvasDataMart();
            //f.Show();
        }

        private void contratosToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            forms.contratacion.gas.FrmContratos f = new forms.contratacion.gas.FrmContratos();
            f.Show();
        }

        private void informeEstadoPuntosÁgoraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmAgora f = new forms.facturacion.FrmAgora();
            f.Show();
        }

        private void reglasOutlookToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.aux1.FrmReglasOutlook f = new forms.aux1.FrmReglasOutlook();
            //f.Show();
        }

        private void sofisticadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmSofisticados f = new forms.facturacion.FrmSofisticados();
            f.Show();
        }

        private void reimpresiónDeContratosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.contratacion.eexxi.FrmReimpresionContratos f = new forms.contratacion.eexxi.FrmReimpresionContratos();
            //f.Show();
        }

        private void importaciónEEXXIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.contratacion.FrmEEXXI f = new forms.contratacion.FrmEEXXI();
            //f.Show();
        }

        private void transponerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmTransponer f = new forms.FrmTransponer();
            //f.Show();
        }

        private void mES13FactoringToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes13Conversion f = new forms.facturacion.FrmMes13Conversion();
            f.ShowDialog();
        }

        private void informeDeBúsquedaDeArchivosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.herramientas.FrmBusquedaArchivos f = new forms.herramientas.FrmBusquedaArchivos();
            f.ShowDialog();
        }

        private void eEXXIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp = new EndesaBusiness.utilidades.Seguimiento_Procesos();

            if (ss_pp.GetFecha_FinProceso("Contratación", "PS_AT", "PS_AT") > ss_pp.GetFecha_InicioProceso("Contratación", "PS_AT", "PS_AT"))
            {
                forms.contratacion.FrmEndesaEnergia21 f = new forms.contratacion.FrmEndesaEnergia21();
                f.ShowDialog();
            }
            else
            {
                MessageBox.Show("El formulario no se puede iniciar debido a que el proceso de actualización todavía no ha terminado."
                    + Environment.NewLine
                    + Environment.NewLine
                    + "Por favor, inténtelo más tarde.",
                   "Endesa Energía XXI",
                   MessageBoxButtons.OK,
                 MessageBoxIcon.Information);
            }
                

            
            
        }

        private void perfilarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.facturacion.FrmPerfilar f = new forms.facturacion.FrmPerfilar();
            //f.Show();
        }

        private void facturaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void festivosEléctricosPortugalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.FrmFacFestivosElectricosParam f = new forms.FrmFacFestivosElectricosParam();
            //f.tabla = "fact.fact_pt_festivos";
            //f.Text = "Festivos Eléctricos Portugal";
            //f.Show();
        }

        private void volcadoCálculoTAMAAccessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //GO.utilidades.Param param = new GO.utilidades.Param("tam_param", MySQLDB.Esquemas.FAC);
            //GO.utilidades.Fichero.EjecutaComando(param.GetValue("Operaciones", DateTime.Now, DateTime.Now), "tamAccess");

        }

        private void exportaciónDeCalendariosTarifariosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.herramientas.FrmCalendarios f = new forms.herramientas.FrmCalendarios();
            f.Show();

            
        }

        private void gasAPSolapadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.gas.FrmInformeSolapados f = new forms.contratacion.gas.FrmInformeSolapados();
            f.Show();
        }

        private void excesosDePotenciaYReactivaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //GO.forms.facturacion.FrmExcesosPotenciaReactiva f = new forms.facturacion.FrmExcesosPotenciaReactiva();
            //f.Show();
        }

        private void ReportesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.medida.FrmKee f = new forms.medida.FrmKee();
            //f.Show();
            forms.medida.FrmKeeReporteExtraccion f = new forms.medida.FrmKeeReporteExtraccion();
            f.Show();
        }

        private void portugalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Environment.MachineName,                       
                    "Facturador Portugal",
                    MessageBoxButtons.OK,
                  MessageBoxIcon.Information);
            
            forms.facturacion.FrmFacturadorPortugal f = new forms.facturacion.FrmFacturadorPortugal();
            f.Show();
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            //forms.utilidades.FrmVarios f = new forms.utilidades.FrmVarios();
            //f.Show();
        }

        private void peticionesCurvasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.medida.Frm_PeticionesCurvas f = new forms.medida.Frm_PeticionesCurvas();
            //f.Show();
        }

        private void importadorXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.eer.FrmImportadorXML f = new forms.eer.FrmImportadorXML();
            //f.Show();
        }

        private void endesaEnergíaRenovableToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void kEEExtracciónFórmulasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.medida.Frm_KEE f = new forms.medida.Frm_KEE();
            //f.Show();
        }

        private void informeParaFiscalGasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmIHGasEspana f = new forms.facturacion.FrmIHGasEspana();
            f.Show();
        }

        private void FacturadorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmFacturadorEER f = new forms.facturacion.FrmFacturadorEER();
            f.Show();
        }

        private void globalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //forms.medida.FrmMedidaGlobal f = new forms.medida.FrmMedidaGlobal();
            //f.Show();
        }

        private void motivosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmMotivosRechazo f = new forms.contratacion.FrmMotivosRechazo();
            f.Show();
        }

        private void paramétrizaciónNotasDeCréditoPortugalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Notas de Crédito Portugal";
            p.tabla = "dsi_ncp_param";
            p.esquemaString = "FAC";
            p.Show();
        }

        private void cOVID19ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Informes COVID";
            p.tabla = "covid_param";
            p.esquemaString = "CON";
            p.Show();
        }

        private void parámetrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.cabecera = "Parametrización Contratos GAS";
            p.tabla = "atrgas_param";
            p.esquemaString = "CON";
            p.Show();
        }

        private void pPAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmPPAs f = new forms.medida.FrmPPAs();
            f.Show();
        }

        private void tunelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmTunel f = new forms.facturacion.FrmTunel();
            f.Show();
        }

        private void facturasBTNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmFacturasBTN f = new forms.facturacion.FrmFacturasBTN();
            f.Show();
        }

        private void carteraSalesForceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.herramientas.FrmCarteraSalesForce f = new forms.herramientas.FrmCarteraSalesForce();
            f.Show();
        }

        private void tunelCuadroMandoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmTunelCuadroMando f = new forms.facturacion.FrmTunelCuadroMando();
            f.Show();
        }

        private void solicitudesManualesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            forms.contratacion.gas.FrmSolicitudManual f = new forms.contratacion.gas.FrmSolicitudManual();
            Cursor.Current = Cursors.Default;
            f.Show();
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void modificaciónPeajeCircularToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            forms.contratacion.gas.FrmModificacionPeaje f = new forms.contratacion.gas.FrmModificacionPeaje();            
            f.Show();
        }

        private void exportaciónTablaPMMLAFTPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmExportacion_PMML f = new forms.medida.FrmExportacion_PMML();
            f.Show();
        }

        private void cuadreDePotenciasMTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmCuadresPotenciaMT f = new forms.facturacion.FrmCuadresPotenciaMT();
            f.Show();
        }

        private void CargaParametros()
        {
            string line;
            int pos;
            string key;
            string value;

            FileInfo file = new FileInfo(System.Environment.CurrentDirectory + @"\bin\" + "properties");
            if (!file.Exists)
                MessageBox.Show("No se encuentra el archivo " + System.Environment.CurrentDirectory + @"\bin\" + "properties",
                    "MySQLDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                System.IO.StreamReader archivo = new System.IO.StreamReader(File.OpenRead(file.FullName));
                while ((line = archivo.ReadLine()) != null)
                {
                    if (line.Trim() != "")
                    {
                        pos = line.IndexOf('=');
                        key = line.Substring(0, pos);
                        value = line.Substring(pos + 1);
                        string a;
                        if (!lista_parametros.TryGetValue(key, out a))
                            lista_parametros.Add(key, value);
                    }

                }
                archivo.Close();
            }
        }

        private string GetValue(string key)
        {
            return lista_parametros.First(z => z.Key == key).Value;
        }

        private void facturasGasPortugalREALESTIMADOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmFacturasGAS_PT_REAL_ESTIMADO f = new forms.facturacion.FrmFacturasGAS_PT_REAL_ESTIMADO();
            f.Show();
        }

        private void pdteWebPSATTAMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "dt_vw_ed_f_detalle_pendiente_facturar_param";
            p.esquemaString = "MED";
            p.cabecera = "Parámetros Pdte web + PS_AT + TAM";
            p.Show();
        }

        private void contrataciónPortugalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "cp_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros Contratación Portugal";
            p.Show();
        }

        private void gestiónPropiaATRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "gestionpropiaatr_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros Gestión Propia ATR";
            p.Show();
        }

        private void informeRevisiónFacturasIRFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "irf_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros Informe Revisión Facturas (IRF)";
            p.Show();
        }

        private void solicitudesATRToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "solatrmt_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros Solicitudes ATR MT";
            p.Show();
        }

        private void pSATToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "ps_param";
            p.esquemaString = "CON";
            p.cabecera = "Parámetros PS_AT";
            p.Show();
        }

        private void generaXMLT105AntiguoAPartirDeA302YA305ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EndesaBusiness.xml.XMLFunciones xml_t105 = new EndesaBusiness.xml.XMLFunciones();
            xml_t105.CargaXML_A302_A305();
        }

        private void pdteWebPSATTAMToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmInforme_Pdte_Web_PSAT_TAM f = new forms.facturacion.FrmInforme_Pdte_Web_PSAT_TAM();
            f.Show();
        }

        private void facturasElectricidadPortugalBTNREALESTIMADOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmFacturasElectricidad_BTN_Real_Estimado f =
                new forms.facturacion.FrmFacturasElectricidad_BTN_Real_Estimado();
            f.Show();
        }

        private void consumoPorMesPortugalDesde2021ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmInformeConsumoMensualPortugal f =
                new forms.facturacion.FrmInformeConsumoMensualPortugal();
            f.Show();
        }

        private void actualizaFacturadoresSofisticadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmActualizaFacturadores f =
                new forms.facturacion.FrmActualizaFacturadores();
            f.Show();
        }

        private void factuarsGasEspañaREALESTIMADOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmFacturasGAS_ES_REAL_ESTIMADO f = new forms.facturacion.FrmFacturasGAS_ES_REAL_ESTIMADO();
            f.Show();
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            //forms.herramientas.
        }

        private void eEXXIBajasSinAltasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmEndesaEnergia21_bajas_sin_altas f = new forms.contratacion.FrmEndesaEnergia21_bajas_sin_altas();
            f.Show();
        }

        private void cálculoPrefacturasBTNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmCalculoPuntosBTN f = new forms.facturacion.FrmCalculoPuntosBTN();
            f.Show();
        }

        private void datosContrataciónXXIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmComprobacion_Contratacion f = new forms.contratacion.FrmComprobacion_Contratacion();
            f.Show();
        }

        private void consumptionDataExtractionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmConsumptionDataExtraction f = new forms.facturacion.FrmConsumptionDataExtraction();
            f.Show();
        }

        private void facturasBTECAPBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmFacturasBTE_CAPB f = new forms.facturacion.FrmFacturasBTE_CAPB();
            f.Show();
        }

        private void prefacturasBTNCAPDeGASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmPrefacturas_BTN_CAP_GAS f = new forms.facturacion.FrmPrefacturas_BTN_CAP_GAS();
            f.Show();
        }

        private void facturasBTNCAPBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmFacturasBTN_CAPB f = new forms.facturacion.FrmFacturasBTN_CAPB();
            f.Show();
        }

        private void tipologíasPortugalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmTipologiasPortugal f = new forms.facturacion.FrmTipologiasPortugal();
            f.Show();
        }

        private void reubicacionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            forms.contratacion.gas.FrmReubicaciones f = new forms.contratacion.gas.FrmReubicaciones();
            f.Show();
            Cursor.Current = Cursors.Default;
        }
        
        private void generadorPasosXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmPasosXML f = new forms.contratacion.FrmPasosXML();
            f.Show();
        }

        private void informesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void informeAnualDGEGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmInformeDGEG f = new forms.medida.FrmInformeDGEG();
            f.Show();
        }

        private void controlProcesosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.herramientas.FrmSeguimiento_Procesos f = new forms.herramientas.FrmSeguimiento_Procesos();
            f.Show();

        }

        private void inventarioTipologíasEspañaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmTipologiasEspana f = new forms.facturacion.FrmTipologiasEspana();
            f.Show();
        }

        private void generadorRespuetasDesdeFicherosXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmXML_Respuestas f = new forms.contratacion.FrmXML_Respuestas();
            f.Show();
        }

        private void factoringMes12BTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes12BT f = new forms.facturacion.FrmMes12BT();
            f.Show();
        }

        private void inventarioBIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmInventarioBI f = new forms.medida.FrmInventarioBI();
            f.Show();
        }

        private void peajesFormatoFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmPeajesFormatoF f = new forms.medida.FrmPeajesFormatoF();
            f.Show();
        }

        private void licitacionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmLicitaciones f = new forms.medida.FrmLicitaciones();
            f.Show();
        }

        private void impagosCAPGASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmImpagosCAP_GAS f = new forms.facturacion.FrmImpagosCAP_GAS();
            f.Show();
        }

        private void generarXMLDesdePlantillaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmPlantillaXML f = new forms.contratacion.FrmPlantillaXML();
            f.Show();
        }

        private void facturaciónADIFToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void importadorXMLA1550ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmXML_A1550 f =
                new forms.contratacion.FrmXML_A1550();
            f.Show();
        }

        private void importadorXMLA1550ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            forms.contratacion.FrmXML_A1550 f =
                new forms.contratacion.FrmXML_A1550();
            f.Show();
        }

        private void formularioAntiguoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdif f = new forms.medida.FrmAdif();
            f.Show();
        }

        private void inventarioToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdifInventario f =
                new forms.medida.FrmAdifInventario();
            f.Show();
        }

        private void análisisMedidaADIFVsBIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdif_CS f = new forms.medida.FrmAdif_CS();
            f.Show();   
        }

        private void exportaciónMedidaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdif_Curvas_Adif f = new forms.medida.FrmAdif_Curvas_Adif();
            f.Show();
        }

        private void facturaciónToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmAdifFacturas f =
              new forms.facturacion.FrmAdifFacturas();
            f.Show();
        }

        private void cierresDeEnergíaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmAdifCierresEnergia f = new forms.medida.FrmAdifCierresEnergia();
            f.Show();
        }

        private void mES13PrevisiónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes13Prevision f = new forms.facturacion.FrmMes13Prevision();
            f.Show();
        }

        private void listaNegraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmMes13ListaNegra f = new forms.facturacion.FrmMes13ListaNegra();
            f.Show();
        }

        private void subestadosPendienteBIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmPendFacturacionSubestadosBI f = new forms.facturacion.FrmPendFacturacionSubestadosBI();
            f.Show();
        }

        private void sAPPdteWebPSATPortugalMTTAMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.FrmParameters p = new forms.FrmParameters();
            p.tabla = "t_ed_h_sap_pendiente_param";
            p.esquemaString = "FAC";
            p.cabecera = "Parámetros Pendiente SAP";
            p.Show();
        }

        private void subestadosPendienteFacturaciónMedidaBIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.facturacion.FrmPendFacturacionSubestadosBI f = new forms.facturacion.FrmPendFacturacionSubestadosBI();
            f.Show();
        }

        private void estadosYSubestadosKronosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forms.medida.FrmEstadosSubEstadosKronos f = new forms.medida.FrmEstadosSubEstadosKronos();
            f.Show();
        }
    }
}
