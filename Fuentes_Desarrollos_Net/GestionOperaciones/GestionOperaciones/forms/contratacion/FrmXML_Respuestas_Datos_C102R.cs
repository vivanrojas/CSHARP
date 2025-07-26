using EndesaBusiness.cnmc;
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
       

    public partial class FrmXML_Respuestas_Datos_C102R : Form
    {
        public EndesaBusiness.cnmc.CNMC cnmc { get; set; }
        //   public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml { get; set; }
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml { get; set; }

        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_A xml_a { get; set; }

        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_R xml_r { get; set; }        //   
        public EndesaBusiness.xml.ContratacionXML cont_xml { get; set; }
    //    public Dictionary<string, string> dic_tipo_activacion_prevista { get; set; }
        public List<string> lista_motivo_rechazo { get; set; }
        
        public DateTime fecha { get; set; }

        public FrmXML_Respuestas_Datos_C102R()
        {
            InitializeComponent();
            lista_motivo_rechazo = new List<string>();

        }

        public FrmXML_Respuestas_Datos_C102R(
                          TipoMensajeC101 xml,
                           EndesaBusiness.cnmc.CNMC cnmc,
                           EndesaBusiness.xml.ContratacionXML cont_xml,
                           List<string> lista_motivo_rechazo)
                          : this()
        {
            this.xml = xml;
            this.cnmc = cnmc;
            this.cont_xml = cont_xml;
            this.lista_motivo_rechazo = lista_motivo_rechazo;
        }

        private void btnOK_Click(object sender, EventArgs e)
         {
            xml_r = new TipoMensajeC102_R();
            // CABECERA → Se construye intercambiando emisora y destino
            xml_r.Cabecera.CodigoREEEmpresaEmisora = xml.Cabecera.CodigoREEEmpresaDestino;
            xml_r.Cabecera.CodigoREEEmpresaDestino = xml.Cabecera.CodigoREEEmpresaEmisora;
            xml_r.Cabecera.CodigoDelProceso = xml.Cabecera.CodigoDelProceso;
            xml_r.Cabecera.CodigoDePaso = "02";
            xml_r.Cabecera.CodigoDeSolicitud = xml.Cabecera.CodigoDeSolicitud;
            xml_r.Cabecera.SecuencialDeSolicitud = xml.Cabecera.SecuencialDeSolicitud;
            xml_r.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            xml_r.Cabecera.CUPS = xml.Cabecera.CUPS;

            //
            //Por ahora solo incluimos un motivo de rechazo
            Rechazo rechazo = new Rechazo();
            rechazo.Secuencial = "01";
            rechazo.CodigoMotivo = cnmc.GetCodigo_Motivo_Rechazo(cmbMotivoRechazo.Text);
            rechazo.Comentarios = textComentario.Text;
            xml_r.Rechazos.Rechazo.Add(rechazo);
            xml_r.Rechazos.FechaRechazo = Convert.ToDateTime(maskedTextFechaRechazo.Text).ToString("yyyy-MM-dd");

            cont_xml.CreaMensajeC102_R(xml_r);
           // cont_xml.CreaMensajeC102_R(c102r);
            

            this.DialogResult = DialogResult.OK;
            this.Close();

           

        }

               

        private void FrmXML_Respuestas_Fecha_Load(object sender, EventArgs e)
        {
            txtboxCUPS.Text = xml?.Cabecera?.CUPS ?? "";
            txtboxCodigoSolicitud.Text = xml?.Cabecera?.CodigoDeSolicitud ?? "";

            cmbMotivoRechazo.Items.Clear();

            if (lista_motivo_rechazo != null)
            {
                foreach (string p in lista_motivo_rechazo)
                {
                    cmbMotivoRechazo.Items.Add(p);
                }
            }
            //if (lista_motivo_rechazo != null)
            //{
            //    for (int i = cmbMotivoRechazo.Items.Count - 1; i >= 0; i--)
            //    {
            //        cmbMotivoRechazo.Items.RemoveAt(i);
            //    }

            //    foreach (string p in lista_motivo_rechazo)
            //    {
            //        cmbMotivoRechazo.Items.Add(p);
            //    }

            //}
            // chk_aceptacion.Checked = true;
            // chk_actuacion_campo.Checked = true;
            //txtboxCUPS.Text = xml.Cabecera.CUPS;
            //txtboxCodigoSolicitud.Text = xml.Cabecera.CodigoDeSolicitud; 
           
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

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void cmbMotivoRechazo_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
