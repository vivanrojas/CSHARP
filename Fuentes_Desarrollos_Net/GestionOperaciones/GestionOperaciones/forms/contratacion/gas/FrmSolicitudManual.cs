using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionOperaciones.forms.contratacion.gas
{
    public partial class FrmSolicitudManual : Form
    {

        
        EndesaBusiness.contratacion.gestionATRGas.GestionATRGas gestionATR;
        EndesaBusiness.cnmc.CNMC cnmc = new EndesaBusiness.cnmc.CNMC();
        EndesaBusiness.cnmc.XML formato_xml = new EndesaBusiness.cnmc.XML();
        EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();
        EndesaBusiness.contratacion.gestionATRGas.Distribuidoras distribuidoras =
            new EndesaBusiness.contratacion.gestionATRGas.Distribuidoras(true);
        EndesaBusiness.sigame.SIGAME inventario_gas = new EndesaBusiness.sigame.SIGAME();
        EndesaBusiness.utilidades.Param pp;

        EndesaBusiness.sigame.Addendas addendas = new EndesaBusiness.sigame.Addendas();

        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();
        public FrmSolicitudManual()
        {

            usage.Start("Contratación", "FrmSolicitudManual" ,"N/A");
            InitializeComponent();
            gestionATR = new EndesaBusiness.contratacion.gestionATRGas.GestionATRGas();
            pp = new EndesaBusiness.utilidades.Param("atrgas_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
        }

        private void FrmSolicitudManual_Load(object sender, EventArgs e)
        {
            cmbTipoProducto.SelectedIndex = 2;
            cmb_tipo.Text = "DIARIO";
            
            txt_fecha_inicio.Value = DateTime.Now.AddDays(1);
            txt_fecha_fin.Value = DateTime.Now.AddDays(1);
            txt_qi.Enabled = false;
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
                
            DateTime temp;
            bool todoOK = true;


            if(chkSoloXML.Checked)
                MessageBox.Show("La solicitud únicamente se va a generar en XML,"
                    + " debido a que está activado (No enviar al SCTD)",
                   "No enviar al SCTD (ACTIVADO)",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);


            if (txt_tipotratamitacion.Text == null || txt_tipotratamitacion.Text == "")
            {
                errorProvider.SetError(txt_cups20, "El CUPS no existe en el inventario.");
                todoOK = false;
            }


            if (txt_cups20.Text == null || txt_cups20.Text == "") 
            {
                errorProvider.SetError(txt_cups20, "El campo debe estar informado.");
                todoOK = false;
            }

            if (cmb_tipo.Text == null || cmb_tipo.Text == "")
            {
                errorProvider.SetError(cmb_tipo, "El campo debe estar informado.");
                todoOK = false;
            }

            if (!DateTime.TryParse(txt_fecha_inicio.Text, out temp))
            {
                errorProvider.SetError(txt_fecha_inicio, "Fecha no válida.");
                todoOK = false;
            }

            if (!DateTime.TryParse(txt_fecha_fin.Text, out temp))
            {
                errorProvider.SetError(txt_fecha_fin, "Fecha no válida.");
                todoOK = false;
            }
            else
            {
                if(txt_fecha_fin.Value < txt_fecha_inicio.Value)
                {
                    errorProvider.SetError(txt_fecha_fin, "Fecha no válida.");
                    todoOK = false;
                }
            }

            if (txt_qd.Enabled && txt_qd.Text == "") 
            {
                errorProvider.SetError(txt_qd, "El campo debe estar informado.");
                todoOK = false;
            }

            if((cmbTipoProducto.SelectedIndex == 0 || cmbTipoProducto.SelectedIndex == 1)
                && txtCodIndProd.Text.Trim() == "")
            {
                errorProvider.SetError(txtCodIndProd, "El campo debe estar informado.");
                todoOK = false;
            }

            if(txt_qd.Enabled && txt_qd.Text == "")
            {
                errorProvider.SetError(txt_qd, "El campo debe estar informado.");
                todoOK = false;
            }

            if (txt_qi.Enabled && txt_qi.Text == "")            
            {                
                errorProvider.SetError(txt_qi, "El campo debe estar informado.");
                todoOK = false;
            }
            
            if (todoOK)
            {
                string textoDarios = "";

                if(cmb_tipo.Text == "DIARIO" && 
                    txt_fecha_fin.Value > txt_fecha_inicio.Value)
                {
                    textoDarios = System.Environment.NewLine
                        + Convert.ToInt32((txt_fecha_fin.Value - txt_fecha_inicio.Value).TotalDays + 1)
                        + " diarios.";
                }
                errorProvider.Clear();
                DialogResult result = MessageBox.Show("¿Desea generar la siguiente solicitud?"
                    + System.Environment.NewLine
                    + System.Environment.NewLine
                    + "Cliente: " + txt_cliente.Text
                    + System.Environment.NewLine
                    + "CUPS20: " + txt_cups20.Text
                    + System.Environment.NewLine
                    + "Producto: " + cmb_tipo.Text
                    + System.Environment.NewLine
                    + "Fecha Desde: " + txt_fecha_inicio.Text
                    + System.Environment.NewLine
                    + "Fecha Desde: " + txt_fecha_fin.Text
                    + System.Environment.NewLine
                    + "Qd: " + string.Format("{0:#,##0}", Convert.ToInt32(txt_qd.Text)) + " kWh/día"                    
                    + textoDarios
                    , "Generar solicitud manual",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    EndesaEntity.contratacion.gas.Solicitud solicitud =
                    new EndesaEntity.contratacion.gas.Solicitud();

                    solicitud.cups = txt_cups20.Text;
                    solicitud.nif = txt_nif.Text;
                    solicitud.razon_social = txt_cliente.Text;
                    solicitud.distribuidora_tramitacion = 
                        distribuidoras.GetTramitacion(inventario_gas.Distribuidora(solicitud.cups), solicitud.nif, solicitud.cups);
                    solicitud.grupo_tarifario = txt_tarifa.Text;

                    if (cmb_tipo.Text == "DIARIO" &&
                        txt_fecha_fin.Value > txt_fecha_inicio.Value)
                    {
                        for (DateTime d = txt_fecha_inicio.Value;
                            d <= txt_fecha_fin.Value; d = d.AddDays(1))
                        {
                            EndesaEntity.contratacion.gas.SolicitudDetalle sd =
                                new EndesaEntity.contratacion.gas.SolicitudDetalle();

                            sd.producto = cmb_tipo.Text;
                            sd.fecha_inicio = d;
                            sd.fecha_fin = d;

                            if(txt_qd.Text != null && txt_qd.Text != "" )
                                sd.qd = Convert.ToDouble(txt_qd.Text);
                            
                            
                            sd.tarifa = txt_tarifa.Text;
                            sd.tipo_producto = cmbTipoProducto.SelectedIndex;
                            if (cmbTipoProducto.SelectedIndex == 0 || 
                                cmbTipoProducto.SelectedIndex == 1)
                            {
                                sd.codigo_producto = txtCodIndProd.Text;
                                sd.tipo_producto = cmbTipoProducto.SelectedIndex;
                            }
                            solicitud.detalle.Add(sd);
                        }
                    }
                    else
                    {
                        EndesaEntity.contratacion.gas.SolicitudDetalle sd =
                            new EndesaEntity.contratacion.gas.SolicitudDetalle();

                        sd.producto = cmb_tipo.Text;
                        sd.fecha_inicio = txt_fecha_inicio.Value;
                        sd.fecha_fin = txt_fecha_fin.Value;
                        sd.qd = Convert.ToDouble(txt_qd.Text);
                        sd.tipo_producto = cmbTipoProducto.SelectedIndex;
                        sd.tarifa = txt_tarifa.Text;
                        if (cmbTipoProducto.SelectedIndex == 0 ||
                                cmbTipoProducto.SelectedIndex == 1)
                        {
                            sd.codigo_producto = txtCodIndProd.Text;
                            sd.tipo_producto = cmbTipoProducto.SelectedIndex;
                        }

                        if(cmb_tipo.SelectedIndex == 3)
                        {
                            if (txt_qi.Text != null && txt_qi.Text != "")
                                sd.qi = Convert.ToDouble(txt_qi.Text);

                            string[] hora = txt_horaInicio.Text.Split(':');
                            sd.hora_inicio = 
                                sd.fecha_inicio.Date.AddHours(Convert.ToInt32(hora[0])).AddMinutes(Convert.ToInt32(hora[1]));
                        }

                        solicitud.detalle.Add(sd);
                    }

                    GeneraXML(solicitud);
                }

                                    
            }


            
        }

        private void GeneraXML(EndesaEntity.contratacion.gas.Solicitud sol)
        {
            EndesaBusiness.cnmc.CNMC cnmc = new EndesaBusiness.cnmc.CNMC();
            EndesaBusiness.cnmc.XML formato_xml = new EndesaBusiness.cnmc.XML();
            //EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();            
            EndesaBusiness.utilidades.ZipUnZip zip = new EndesaBusiness.utilidades.ZipUnZip();

            string secuencial;
            string fileName = "";
            string fechaHora = "";
            string destinycompany = "";

            try
            {

                if (!Directory.Exists(pp.GetValue("inbox", DateTime.Now, DateTime.Now)))
                    Directory.CreateDirectory(pp.GetValue("inbox", DateTime.Now, DateTime.Now));

                BorrarContenidoDirectorio(pp.GetValue("inbox", DateTime.Now, DateTime.Now));

                //secuencial = Convert.ToInt32(pp.GetValue("secuencial_solicitud", DateTime.Now, DateTime.Now));

                for (int i = 0; i < sol.detalle.Count(); i++)
                {
                    Thread.Sleep(1000);
                    secuencial = DateTime.Now.ToString("yyMMddHHmmss").ToString();
                    gestionATR.GuardaNumSecuencialTemporal(secuencial);
                    Thread.Sleep(1000);

                    EndesaEntity.cnmc.XML_A1_43 xml_a1_43 = new EndesaEntity.cnmc.XML_A1_43();

                    //xml_a1_43.comreferencenum = pp.GetValue("prefijo_solicitud", DateTime.Now, DateTime.Now)
                    //        + DateTime.Now.ToString("yyyy") + secuencial.ToString().PadLeft(4, '0');

                    xml_a1_43.comreferencenum = secuencial;

                    xml_a1_43.destinycompany =
                            distribuidoras.Codigo_XML_CNMC_Distribuidora(inventario_gas.Distribuidora(sol.cups).ToUpper());
                    destinycompany = xml_a1_43.destinycompany;
                    xml_a1_43.dispatchingcompany = "0007";
                    xml_a1_43.documentnum = inventario_gas.NIF(sol.cups);
                    xml_a1_43.cups = sol.cups;
                    xml_a1_43.productstartdate = sol.detalle[i].fecha_inicio;
                    xml_a1_43.reqtype = sol.detalle[i].tipo_producto;
                    xml_a1_43.productcode = sol.detalle[i].codigo_producto;
                    xml_a1_43.producttype = cnmc.Codigo_Tipo_Producto(sol.detalle[i].producto.ToUpper());
                    //xml_a1_43.producttolltype = cnmc.Codigo_Tipo_Peaje(inventario_gas.Grupo_Presion(sol.cups),
                    //    sol.detalle[i].qd * 330);
                    xml_a1_43.producttolltype = cnmc.Codigo_Tipo_Peaje(sol.detalle[i].tarifa);



                    xml_a1_43.productqd = sol.detalle[i].qd;
                    xml_a1_43.productqi = sol.detalle[i].qi;
                    xml_a1_43.startHour = sol.detalle[i].hora_inicio;

                    fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                    fileName = pp.GetValue("inbox", DateTime.Now, DateTime.Now)
                        + @"\" + pp.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                        + xml_a1_43.destinycompany + "_"
                        + fechaHora
                        + ".xml";
                    FileInfo file = new FileInfo(fileName);

                    formato_xml.CreaXML_A1_43(file, xml_a1_43);
                    

                }

                // Comprimimos los XML y los enviamos por mail
                fechaHora = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                fileName = pp.GetValue("inbox", DateTime.Now, DateTime.Now) + @"\"
                        + pp.GetValue("prefijo_archivo_a1_43", DateTime.Now, DateTime.Now)
                        + destinycompany + "_"
                        + fechaHora
                        + ".zip";
                FileInfo archivo = new FileInfo(fileName);

                //zip.ComprimirVarios(pp.GetValue("inbox", DateTime.Now, DateTime.Now) + @"\",
                //    ".*\\.(xml)$", archivo.FullName);

                zip.ComprimirVarios(pp.GetValue("inbox") + @"\" + "*.xml", archivo.FullName);


                if (!chkSoloXML.Checked)
                {
                    if (sol.distribuidora_tramitacion == "XML_SCTD")
                        gestionATR.Subir_a_FTP();
                    else if (sol.distribuidora_tramitacion == "XML")
                        gestionATR.Subir_a_FTP_Extremadura();
                }
               
                

                //if (pp.GetValue("enviar_mail_XML", DateTime.Now, DateTime.Now) == "S")
                //    GeneraMail(fileName);

                if (chkSolapado.Checked)
                    gestionATR.GuardaSolicitud(sol);

                if (!gestionATR.error_ftp)
                {
                    MessageBox.Show("Solicitud Generada correctamente en"
                        + " " + pp.GetValue("inbox"),
                      "Solicitud Completada",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                    this.Close();
                }
                else
                {
                    MessageBox.Show("La solicitud no se ha podido completar.",
                    "Error al generar la solicitud",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                }
                    

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error al generar la solicitud",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        private void txt_cups20_TextChanged(object sender, EventArgs e)
        {
            EndesaBusiness.utilidades.UltimateFTP ftp;
            string tarifa = null;

            if (txt_cups20.Text.Trim().Length == 20)
            {
                txt_nif.Text = inventario_gas.NIF(txt_cups20.Text.Trim());
                txt_cliente.Text = inventario_gas.NombreCliente(txt_cups20.Text.Trim());
                txt_distribuidora.Text = inventario_gas.Distribuidora(txt_cups20.Text.Trim());
                txt_tipotratamitacion.Text = distribuidoras.GetTramitacion(txt_distribuidora.Text,txt_nif.Text, txt_cups20.Text.Trim());
                //txt_tarifa.Text = inventario_gas.tarifa(txt_cups20.Text.Trim()):
                tarifa = addendas.UltimaTarifaAddendas(txt_cups20.Text.Trim());
                if (tarifa != null)
                    txt_tarifa.Text = tarifa;
                else
                    txt_tarifa.Text = inventario_gas.Tarifa(txt_cups20.Text.Trim());

                if (txt_nif.Text == "" || txt_distribuidora.Text == "")
                {
                    errorProvider.SetError(txt_cups20, "CUPS desconocido");
                    
                }



                if(txt_tipotratamitacion.Text == "XML")
                {
                    
                    ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        pp.GetValue("ftp_extremadura_server", DateTime.Now, DateTime.Now),
                        pp.GetValue("ftp_extremadura_user", DateTime.Now, DateTime.Now),
                        EndesaBusiness.utilidades.FuncionesTexto.Decrypt(pp.GetValue("ftp_extremadura_pass", DateTime.Now, DateTime.Now), true),
                        pp.GetValue("ftp_extremadura_port", DateTime.Now, DateTime.Now));

                    if (!ftp.isValidConnectionFTPS())
                    {
                        MessageBox.Show("No tiene acceso al FTP parametrizado y no puede realizar la solicitud.",
                           "Error de aceso al FTP",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error);

                        chkSoloXML.Checked = true;
                    }
                }else if (txt_tipotratamitacion.Text == "XML_SCTD")
                {
                    ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        pp.GetValue("ftp_sctd_server", DateTime.Now, DateTime.Now),
                        pp.GetValue("ftp_sctd_user", DateTime.Now, DateTime.Now),
                        EndesaBusiness.utilidades.FuncionesTexto.Decrypt(pp.GetValue("ftp_sctd_pass", DateTime.Now, DateTime.Now), true),
                        pp.GetValue("ftp_sctd_port", DateTime.Now, DateTime.Now));
                }

            }
            else
            {                
                errorProvider.SetError(txt_cups20, "El CUPS no tiene 20 caracteres");                
            }
            
            
        }

        private void txt_qd_TextChanged(object sender, EventArgs e)
        {
            if (txt_qd.Text != "")
            {
                txt_qi.Text = "";
                txt_qi.Enabled = false;
            }
            else
            {
                //txt_tarifa.Text =
                //    cnmc.Tarifa(inventario_gas.Grupo_Presion(txt_cups20.Text.Trim()),
                //      Convert.ToDouble(txt_qd.Text) * 330);

                txt_qi.Enabled = true;


            }
               
        }


        private void txt_qi_TextChanged(object sender, EventArgs e)
        {
            if (txt_qi.Text != "")
            {
                txt_qd.Text = "";
                txt_qd.Enabled = false;
            }
            else
            {


                txt_qd.Enabled = true;


            }
        }

        private void txt_tarifa_TextChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmb_tipo_SelectedIndexChanged(object sender, EventArgs e)
        {
            ValidaFechas();
        }

        private void txt_qd_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                              
                e.Handled = true;
            }
          

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
            
        }

        private void ValidaFechas()
        {
            DateTime fd = new DateTime();
            DateTime fh = new DateTime();
            txt_fecha_fin.Enabled = false;

            txt_qi.Text = "";
            txt_horaInicio.Text = "";
            txt_qi.Enabled = false;
            txt_horaInicio.Enabled = false;


            fd = txt_fecha_inicio.Value;

            if (cmb_tipo.Text == "DIARIO")
                txt_fecha_fin.Enabled = true;

            if (cmb_tipo.Text == "MENSUAL")
            {
                if(fd.Month == DateTime.Now.Month)
                    fd = fd.AddMonths(1);
                txt_fecha_inicio.Value = new DateTime(fd.Year, fd.Month, 1);
                txt_fecha_fin.Value =
                    new DateTime(fd.Year, fd.Month, DateTime.DaysInMonth(fd.Year, fd.Month));
            }

            if (cmb_tipo.Text == "TRIMESTRAL")
            {
                if (fd.Month == DateTime.Now.Month)
                    fd = DateTime.Now.AddMonths(1);
                txt_fecha_inicio.Value = new DateTime(fd.Year, fd.Month, 1);
                fh = txt_fecha_inicio.Value.AddMonths(2);
                txt_fecha_fin.Value =
                    new DateTime(fh.Year, fh.Month, DateTime.DaysInMonth(fh.Year, fh.Month));
            }

            if (cmb_tipo.Text == "INDEFINIDO")
            {
                if (fd.Month == DateTime.Now.Month)
                    fd = DateTime.Now.AddMonths(1);
                // txt_fecha_inicio.Value = new DateTime(fd.Year, fd.Month, 1);
                txt_fecha_fin.Value =
                    new DateTime(4999, 12, 31);
            }

            if (cmb_tipo.Text == "ANUAL")
            {
                if (fd.Month == DateTime.Now.Month)
                    fd = DateTime.Now.AddMonths(1);
                txt_fecha_inicio.Value = new DateTime(fd.Year, fd.Month, 1);
                fh = txt_fecha_inicio.Value.AddYears(1);
                txt_fecha_fin.Value =
                    new DateTime(fh.Year, fh.Month -1, DateTime.DaysInMonth(fh.Year, fh.Month -1));
            }
            if (cmb_tipo.Text == "INTRADIARIO")
            {
                txt_qi.Enabled = true;
                txt_horaInicio.Enabled = true;
                txt_horaInicio.Text = DateTime.Now.AddHours(1).ToString("HH:00");
            }

        }

        private void txt_fecha_inicio_ValueChanged(object sender, EventArgs e)
        {
            if (cmb_tipo.Text == "DIARIO")
            {
                txt_fecha_fin.Enabled = false;
                txt_fecha_fin.Value = txt_fecha_inicio.Value;
            }

            if (cmb_tipo.Text == "INTRADIARIO")
            {                
                txt_qi.Enabled = true;
                txt_horaInicio.Enabled = true;
                txt_fecha_fin.Value = txt_fecha_inicio.Value;
            }

            ValidaFechas();
        }

        private void BorrarContenidoDirectorio(string directorio)
        {
            string[] listaArchivos;
            FileInfo file;

            listaArchivos = Directory.GetFiles(directorio);
            for (int i = 0; i < listaArchivos.Count(); i++)
            {
                file = new FileInfo(listaArchivos[i]);
                file.Delete();
            }
        }

        private void txt_qd_Leave(object sender, EventArgs e)
        {
            //txt_tarifa.Text =
            //         cnmc.Tarifa(inventario_gas.Grupo_Presion(txt_cups20.Text.Trim()),
            //           Convert.ToDouble(txt_qd.Text) * 330);
        }

        private void cmbTipoProducto_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtCodIndProd.Enabled = 
                (cmbTipoProducto.SelectedIndex == 0 || cmbTipoProducto.SelectedIndex == 1);
        }

        private void txtCodIndProd_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void toolTip1_Popup_1(object sender, PopupEventArgs e)
        {

        }

        private void txt_qi_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {

                e.Handled = true;
            }


            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void FrmSolicitudManual_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmSolicitudManual" ,"N/A");
        }
    }
}
