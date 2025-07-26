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
using EndesaEntity.cnmc.V30_2022_21_01;

namespace GestionOperaciones.forms.contratacion
{
       

    public partial class FrmXML_Respuestas_Datos_M102 : Form
    {
        public EndesaBusiness.cnmc.CNMC cnmc { get; set; }
      //  public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml { get; set; }
      //  public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA302_A xml_a { get; set; }

        public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml { get; set; }

       public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM102A  xml_a { get; set; }

        public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM102_R xml_r { get; set; }

     //   public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA302_R xml_r { get; set; }
        public EndesaBusiness.xml.ContratacionXML cont_xml { get; set; }
        public Dictionary<string, string> dic_tipo_activacion_prevista { get; set; }
        public List<string> lista_motivo_rechazo { get; set; }
        
        public DateTime fecha { get; set; }

        public FrmXML_Respuestas_Datos_M102()
        {
            InitializeComponent();
            this.AutoValidate = AutoValidate.Disable;
        }
        public FrmXML_Respuestas_Datos_M102(
             EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml,
             EndesaBusiness.cnmc.CNMC cnmc,
             EndesaBusiness.xml.ContratacionXML cont_xml)
             : this() // Llama al constructor sin parámetros para mantener InitializeComponent
             {

            if (xml == null)
                throw new ArgumentNullException(nameof(xml), "El objeto XML (TipoMensajeM101) no puede ser nulo.");

            if (cnmc == null)
                throw new ArgumentNullException(nameof(cnmc), "El objeto CNMC no puede ser nulo.");

            if (cont_xml == null)
                throw new ArgumentNullException(nameof(cont_xml), "El objeto ContratacionXML no puede ser nulo.");

            this.xml = xml;
            this.cnmc = cnmc;
            this.cont_xml = cont_xml;
             }
        // 
        private void btnOK_Click(object sender, EventArgs e)
        {
            DateTime temp;
            bool todoOK = true;
            string actuacion_campo = "";
            bool es_aceptacion = chk_aceptacion.Checked; 
            
            if (es_aceptacion && !DateTime.TryParse(txt_FechaAceptacion.Text, out temp))
            {
                errorProvider.SetError(txt_FechaAceptacion, "Fecha no válida.");
                todoOK = false;
            }
            else if (!es_aceptacion && !DateTime.TryParse(maskedTextFechaRechazo.Text, out temp))
{
                errorProvider.SetError(txt_FechaAceptacion, "Fecha no válida.");
                todoOK = false;
            }

            if (todoOK)
            {
                
                if (es_aceptacion) //Es aceptacion
                {
                    //xml_a = new TipoMensajeA302_A();
                    xml_a = new TipoMensajeM102A();
                    xml_a.Cabecera.CodigoREEEmpresaEmisora = xml.Cabecera.CodigoREEEmpresaDestino;
                    xml_a.Cabecera.CodigoREEEmpresaDestino = xml.Cabecera.CodigoREEEmpresaEmisora;
                    xml_a.Cabecera.CodigoDelProceso = xml.Cabecera.CodigoDelProceso;
                    xml_a.Cabecera.CodigoDePaso = "02";
                    xml_a.Cabecera.CodigoDeSolicitud = xml.Cabecera.CodigoDeSolicitud;
                    xml_a.Cabecera.SecuencialDeSolicitud = xml.Cabecera.SecuencialDeSolicitud;
                    xml_a.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss").ToString();
                    xml_a.Cabecera.CUPS = xml.Cabecera.CUPS;

                   // xml_a.AceptacionAlta.DatosAceptacion.fechaAceptacion = DateTime.Now.ToString("yyyy-MM-dd").ToString();
                   // xml_a.AceptacionAlta.DatosAceptacion.ActuacionCampo = actuacion_campo = chk_actuacion_campo.Checked ? "S" : "N"; 
                    //xml_a.AceptacionAlta.Contrato.CondicionesContractuales.TarifaATR = xml.Alta.Contrato.CondicionesContractuales.TarifaATR;

                    //foreach (Potencia p in xml.Alta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia)
                    //{
                    //    Potencia c = new Potencia();
                    //    c.potencia = p.potencia;
                    //    c.periodo = p.periodo;
                    //    xml_a.AceptacionAlta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia.Add(c);
                    //}

                    
                    //xml_a.AceptacionAlta.Contrato.CondicionesContractuales.ModoControlPotencia = xml.Alta.Contrato.CondicionesContractuales.ModoControlPotencia;
                    //xml_a.AceptacionAlta.Contrato.TipoActivacionPrevista = xml.Alta.Contrato.TipoActivacionPrevista;
                    //xml_a.AceptacionAlta.Contrato.FechaActivacionPrevista = xml.Alta.Contrato.FechaActivacionPrevista;

                    //xml_a.AceptacionAlta.Contrato.TipoActivacionPrevista = cnmc.GetCodigo("cnmc_p_tipo_activacion_prevista", cmb_tipo_activacion.Text);
                    //xml_a.AceptacionAlta.Contrato.FechaActivacionPrevista = Convert.ToDateTime(txt_FechaAceptacion.Text).ToString("yyyy-MM-dd");
                    //xml_a.AceptacionAlta.Contrato.Contacto = null;

                    //cont_xml.CreaMensajeA302_A(xml_a);
                    cont_xml.CreaMensajeM102_A(xml_a);
                }
                else //Es rechazo
                {
                    //xml_r = new TipoMensajeA302_R();
                    xml_r = new TipoMensajeM102_R();
                    xml_r.Cabecera.CodigoREEEmpresaEmisora = xml.Cabecera.CodigoREEEmpresaDestino;
                    xml_r.Cabecera.CodigoREEEmpresaDestino = xml.Cabecera.CodigoREEEmpresaEmisora;
                    xml_r.Cabecera.CodigoDelProceso = xml.Cabecera.CodigoDelProceso;
                    xml_r.Cabecera.CodigoDePaso = "02";
                    xml_r.Cabecera.CodigoDeSolicitud = xml.Cabecera.CodigoDeSolicitud;
                    xml_r.Cabecera.SecuencialDeSolicitud = xml.Cabecera.SecuencialDeSolicitud;
                    xml_r.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss").ToString();
                    xml_r.Cabecera.CUPS = xml.Cabecera.CUPS;

                    //Por ahora solo incluimos un motivo de rechazo
                    Rechazo rechazo = new Rechazo();
                    rechazo.Secuencial = "01";
                    rechazo.CodigoMotivo = cnmc.GetCodigo_Motivo_Rechazo(cmbMotivoRechazo.Text);
                    rechazo.Comentarios = textComentario.Text;
                    xml_r.Rechazos.Rechazo.Add(rechazo);
                    xml_r.Rechazos.FechaRechazo = Convert.ToDateTime(maskedTextFechaRechazo.Text).ToString("yyyy-MM-dd");

                    // cont_xml.CreaMensajeA302_R(xml_r);
                    cont_xml.CreaMensajeM102_R(xml_r);
                }


                //actuacion_campo = chk_actuacion_campo.Checked ? "S" : "N";
                
                
                //cont_xml.CreaMensajeA302(xml, actuacion_campo, es_aceptacion);
                //cont_xml.CreaMensajeA302_R(xml);
                this.Close();
            }


        }

        private void FrmXML_Respuestas_Fecha_Load(object sender, EventArgs e)
        {
            for (int i = cmb_tipo_activacion.Items.Count - 1; i == 0; i--)
            {
                cmb_tipo_activacion.Items.RemoveAt(i);
            }

            foreach (KeyValuePair<string, string> p in dic_tipo_activacion_prevista)
            {
                cmb_tipo_activacion.Items.Add(p.Key);
            }

            for (int i = cmbMotivoRechazo.Items.Count - 1; i == 0; i--)
            {
                cmbMotivoRechazo.Items.RemoveAt(i);
            }

            foreach (string p in lista_motivo_rechazo)
            {
                cmbMotivoRechazo.Items.Add(p);
            }

            chk_aceptacion.Checked = true;
            chk_actuacion_campo.Checked = true;
            txtboxCUPS.Text = xml.Cabecera.CUPS;
            txtboxCodigoSolicitud.Text = xml.Cabecera.CodigoDeSolicitud;
            //cmb_tipo_activacion.SelectedIndex = 3;

        }

        private void chk_rechazo_CheckedChanged(object sender, EventArgs e)
        {
            chk_aceptacion.Checked = !chk_rechazo.Checked;
            groupBox1.Enabled = !chk_rechazo.Checked;
        }

        private void chk_aceptacion_CheckedChanged(object sender, EventArgs e)
        {
            chk_rechazo.Checked = !chk_aceptacion.Checked;
            groupBox2.Enabled = !chk_aceptacion.Checked;
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void cmbMotivoRechazo_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
