using EndesaEntity.medida;
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
using System.Xml.Schema;
using System.Xml;
using System.Xml.Serialization;
using EndesaBusiness.sharePoint;
using static EndesaBusiness.servidores.OracleServer;
using EndesaBusiness.xml;
using static EndesaBusiness.medida.Kee_Extraccion_Formulas;
using EndesaBusiness.cnmc;
using EndesaEntity.cnmc.V30_2022_21_01;
using EndesaEntity.cnmc.V21_2019_12_17;


namespace GestionOperaciones.forms.contratacion
{
    public partial class FrmXML_Respuestas : Form
    {

        EndesaBusiness.cnmc.CNMC cnmc;
        EndesaBusiness.xml.ContratacionXML cont_xml;
        EndesaBusiness.utilidades.UsageForm usage = new EndesaBusiness.utilidades.UsageForm();  //   ojo ver rendimiento ----13/07/25
        public FrmXML_Respuestas()
        {
            usage.Start("Contratación", "FrmXML_Respuestas" ,"N/A");
            InitializeComponent();

            //
            this.AutoValidate = AutoValidate.Disable;

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //private void cargaXMLToolStripMenuItem_Click(object sender, EventArgs e) ---------- es la primera opcion..... solo carga xml de A301 
        //{
        //    DialogResult result = DialogResult.Yes;
        //    EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml;


        //    try
        //    {
        //        OpenFileDialog d = new OpenFileDialog();
        //        d.Title = "Carga XML";
        //        d.Filter = "xml files|*.xml";
        //        d.Multiselect = true;

        //        if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //        {
        //            foreach (string fileName in d.FileNames)
        //            {

        //                XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301));

        //                System.Xml.Serialization.XmlSerializer reader =
        //                    new System.Xml.Serialization.XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301));
        //                System.IO.StreamReader file = new System.IO.StreamReader(fileName);
        //                xml = (EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301)reader.Deserialize(file);
        //                file.Close();

        //                cont_xml.Guarda_XML(xml);

        //            }

        //            dgv.AutoGenerateColumns = false;
        //            dgv.DataSource = cont_xml.GetLista();
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        MessageBox.Show(ex.Message,
        //        "Error en carga XML",
        //        MessageBoxButtons.OK,
        //        MessageBoxIcon.Error);
        //    }

        //}
        //private void cargaXMLToolStripMenuItem_Click(object sender, EventArgs e) //  segunda opcion carga xml de todos los tipos de mensajes
        //{
        //    try
        //    {
        //        OpenFileDialog d = new OpenFileDialog
        //        {
        //            Title = "Carga XML",
        //            Filter = "xml files|*.xml",
        //            Multiselect = true
        //        };

        //        if (d.ShowDialog() == DialogResult.OK)
        //        {
        //            foreach (string fileName in d.FileNames)
        //            {
        //                string xmlContent = File.ReadAllText(fileName);

        //                //if (xmlContent.Contains("<TipoMensajeA301"))
        //                if (xmlContent.Contains("<MensajeAlta"))
        //                {


        //                    // vallidador del xml 
        //                    //var reader = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301));
        //                    //using (var file = new StreamReader(fileName))
        //                    //{
        //                    //    var xml = (EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301)reader.Deserialize(file);
        //                    //  //  cont_xml.Guarda_XML(xml);
        //                    //}
        //                    var reader = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301));
        //                    using (var file = new StreamReader(fileName))
        //                    {
        //                        var xml = (EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301)reader.Deserialize(file);
        //                        cont_xml.Guarda_XMLV30_a301(xml);
        //                    }
        //                }
        //                else// if (xmlContent.Contains("<TipoMensajeC101"))
        //                 if (xmlContent.Contains("<MensajeCambiodeComercializadorSinCambios"))
        //                {
        //                    var reader = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101));
        //                    using (var file = new StreamReader(fileName))
        //                    {
        //                        var xml = (EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101)reader.Deserialize(file);
        //                        cont_xml.Guarda_XMLc101(xml);
        //                    }
        //                }
        //                else if (xmlContent.Contains("<MensajeBajaSuspension"))
        //                {
        //                    var reader = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeB101));
        //                    using (var file = new StreamReader(fileName))
        //                    {
        //                        var xml = (EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeB101)reader.Deserialize(file);
        //                        cont_xml.Guarda_XMLb101(xml);
        //                    }
        //                }
        //                else if (xmlContent.Contains("<MensajeCambiodeComercializadorConCambios"))
        //                {
        //                    var reader = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201));
        //                    using (var file = new StreamReader(fileName))
        //                    {
        //                        var xml = (EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201)reader.Deserialize(file);
        //                        cont_xml.Guarda_XMLc201(xml);
        //                    }
        //                }
        //                else if (xmlContent.Contains("<MensajeModificacionDeATR"))
        //                {
        //                    var reader = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101));
        //                    using (var file = new StreamReader(fileName))
        //                    {
        //                        var xml = (EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101)reader.Deserialize(file);
        //                        cont_xml.Guarda_XMLm101(xml);
        //                    }
        //                }
        //                else
        //                {
        //                    MessageBox.Show($"No se reconoce el tipo de mensaje en el archivo:\n{fileName}", "Tipo no reconocido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                }
        //            }

        //            dgv.AutoGenerateColumns = false;
        //            dgv.DataSource = cont_xml.GetLista(); // Este método debe devolver todos los tipos cargados
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message,
        //            "Error en carga XML",
        //            MessageBoxButtons.OK,
        //            MessageBoxIcon.Error);
        //    }
        //}


        //private void cargaXMLToolStripMenuItem_Click(object sender, EventArgs e)  // tercera opcion carga xml de todos los tipos de mensajes con validacion de xsd
        private void cargaXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog d = new OpenFileDialog
                {
                    Title = "Carga XML",
                    Filter = "xml files|*.xml",
                    Multiselect = true
                };
                // -----------------------ojo  revisar---------------------------04/07/2025
                if (d.ShowDialog() == DialogResult.OK)
                {
                    // ✅ Instanciar param_cnmc UNA SOLA VEZ
                    var param_cnmc = new EndesaBusiness.utilidades.Param("cnmc_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                    // Definir la carpeta BASE donde tienes TODOS tus XSD
                    string xsdFolder = @"C:\CSHARP\Fuentes_Desarrollos_Net\GestionOperaciones\GestionOperaciones\bin\x64\Debug\media\xsd\CNMC - E - XSD 2024.05.16";

                    foreach (string fileName in d.FileNames)
                    {
                        string xmlContent = File.ReadAllText(fileName);

                        // Determinar tipo de mensaje y configuración de validación
                        string mensaje = "";
                        Type tipo = null;
                        Action<object> guardarAccion = null;
                        string xsdParamKey = null;

                        if (xmlContent.Contains("<MensajeAlta"))
                        {
                            mensaje = "A301";
                            tipo = typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301);
                            guardarAccion = obj =>
                                cont_xml.Guarda_XMLV30_a301((EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301)obj);
                            xsdParamKey = "xsd_a301";
                        }
                        else if (xmlContent.Contains("<MensajeCambiodeComercializadorSinCambios"))
                        {
                            mensaje = "C101";
                            tipo = typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101);
                            guardarAccion = obj =>
                                cont_xml.Guarda_XMLc101((EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101)obj);
                            xsdParamKey = "xsd_c101";
                        }
                        else if (xmlContent.Contains("<MensajeBajaSuspension"))
                        {
                            mensaje = "B101";
                            tipo = typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeB101);
                            guardarAccion = obj =>
                                cont_xml.Guarda_XMLb101((EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeB101)obj);
                            xsdParamKey = "xsd_b101";
                        }
                        else if (xmlContent.Contains("<MensajeCambiodeComercializadorConCambios"))
                        {
                            mensaje = "C201";
                            tipo = typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201);
                            guardarAccion = obj =>
                                cont_xml.Guarda_XMLc201((EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201)obj);
                            xsdParamKey = "xsd_c201";
                        }
                        else if (xmlContent.Contains("<MensajeModificacionDeATR"))
                        {
                            mensaje = "M101";
                            tipo = typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101);
                            guardarAccion = obj =>
                                cont_xml.Guarda_XMLm101((EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101)obj);
                            xsdParamKey = "xsd_m101";
                        }
                        else
                        {
                            MessageBox.Show(
                                $"No se reconoce el tipo de mensaje en el archivo:\n{fileName}",
                                "Tipo no reconocido",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            continue;
                        }
                        // Obtener ruta del XSD desde param
                        string rutaXsd = param_cnmc.GetValue(xsdParamKey);

                        if (string.IsNullOrEmpty(rutaXsd))
                        {
                            MessageBox.Show(
                                $"No se encontró la ruta del XSD en la tabla param para {xsdParamKey}.",
                                "XSD no definido",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            continue;
                        }
                        // Extraer solo el nombre del archivo XSD
                        string xsdFileName = Path.GetFileName(rutaXsd);

                        // Construir la ruta absoluta en la carpeta fija
                        string xsdPath = Path.Combine(xsdFolder, xsdFileName);

                        if (!File.Exists(xsdPath))
                        {
                            MessageBox.Show(
                                $"No se encontró el archivo XSD:\n{xsdPath}",
                                "Archivo XSD no encontrado",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            continue;
                        }

                        // ✅ Obtener path del XSD usando param_cnmc
                        //string rutaXsd = param_cnmc.GetValue(xsdParamKey);

                        // ✅ Montar path absoluto
                        //string xsdPath = GetXsdFullPath(rutaXsd);

                        //if (string.IsNullOrEmpty(xsdPath) || !File.Exists(xsdPath))
                        //{
                        //    MessageBox.Show(
                        //        $"No se encontró el archivo XSD:\n{xsdPath}",
                        //        "Archivo XSD no encontrado",
                        //        MessageBoxButtons.OK,
                        //        MessageBoxIcon.Error);
                        //    continue;
                        //}

                        // ✅ Validar XML contra XSD
                        string validationErrors = ValidateXmlAgainstXsd(fileName, xsdPath);

                        if (!string.IsNullOrEmpty(validationErrors))
                        {
                            MessageBox.Show(
                                $"Errores en el XML ({mensaje}) del archivo {Path.GetFileName(fileName)}:\n\n{validationErrors}",
                                "Error de validación",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                            continue;
                        }

                        // Si es válido, deserializar y guardar
                        var serializer = new XmlSerializer(tipo);
                        using (var file = new StreamReader(fileName))
                        {
                            var xmlObj = serializer.Deserialize(file);
                            guardarAccion(xmlObj);
                        }
                    }

                    dgv.AutoGenerateColumns = false;
                    dgv.DataSource = cont_xml.GetLista();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Error en carga XML",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        public static string ValidateXmlAgainstXsd(string xmlPath, string xsdPath)  // irh
        {
            string validationErrors = "";

            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add(null, xsdPath);

            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationType = ValidationType.Schema;
            settings.Schemas = schemas;
            settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            settings.ValidationEventHandler += (sender, e) =>
            {
                validationErrors += e.Message + Environment.NewLine;
            };

            using (XmlReader reader = XmlReader.Create(xmlPath, settings))
            {
                try
                {
                    while (reader.Read()) { }
                }
                catch (Exception ex)
                {
                    validationErrors += ex.Message + Environment.NewLine;
                }
            }

            return validationErrors;
        }

        private string GetXsdFullPath(string rutaXsdDevuelta)  // irh
        {
            if (string.IsNullOrEmpty(rutaXsdDevuelta))
                return null;

            if (Path.IsPathRooted(rutaXsdDevuelta))
            {
                return rutaXsdDevuelta;
            }
            else
            {
                if (rutaXsdDevuelta.StartsWith(@"\"))
                {
                    return @"C:" + rutaXsdDevuelta;
                }
                else
                {
                    return Path.Combine(Environment.CurrentDirectory, rutaXsdDevuelta);
                }
            }
        }


        private void FrmXML_Respuestas_Load(object sender, EventArgs e)
        {
            cnmc = new EndesaBusiness.cnmc.CNMC();
            cont_xml = new EndesaBusiness.xml.ContratacionXML();

            dgv.AutoGenerateColumns = false;
            dgv.DataSource = cont_xml.GetLista();
        }

        // Botón Aceptación A302-A y Rechazo A302-R
        private void btn_a302_Click(object sender, EventArgs e)
        {

            EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml;
           //DateTime fecha_activacion_prevista = new DateTime();

            
            foreach (DataGridViewRow row in dgv.SelectedRows)
            {

                DataGridViewCell id = (DataGridViewCell)
                 row.Cells[0];



                //xml = cont_xml.CargaDatos(Convert.ToInt32(id.Value.ToString()));--------------
                xml = cont_xml.CargaDatosV30(Convert.ToInt32(id.Value.ToString()));
                FrmXML_Respuestas_Datos_A3 f = new FrmXML_Respuestas_Datos_A3();
                f.cnmc = cnmc;
                f.xml = xml;
                f.cont_xml = cont_xml;
                f.dic_tipo_activacion_prevista = cnmc.dic_tipo_activacion_prevista;
                f.lista_motivo_rechazo = cnmc.GetLista_Motivos_Rechazo("A3");
                   
                f.ShowDialog();

                

                MessageBox.Show(@"Mensajes creados en c:\Temp",
               "Creación Paso A302",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information);

            }

           
        }
        // Botón Activación A305
        private void btn_a305_Click(object sender, EventArgs e)
        {
            EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml;
            

            foreach (DataGridViewRow row in dgv.SelectedRows)
            {

                DataGridViewCell id = (DataGridViewCell)
                 row.Cells[0];
                // llamo  ContratacionXML  /
                xml = cont_xml.CargaDatosV30(Convert.ToInt32(id.Value.ToString()));
              // control de campos  a mostrar si no tipoautoconsumo no mostrar sera null.

                FrmXML_Respuestas_Datos_A305 f = new FrmXML_Respuestas_Datos_A305();
                f.cnmc = cnmc;
                f.xml = xml;
                f.cont_xml = cont_xml;
                //f.dic_tipo_activacion_prevista = cnmc.dic_tipo_activacion_prevista;
                //f.lista_motivo_rechazo = cnmc.GetLista_Motivos_Rechazo("A3");
                var result = f.ShowDialog();

                if (result == DialogResult.OK)
                {
                    MessageBox.Show(@"Mensajes creados en c:\Temp",
                        "Creación Paso A305",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

            }

            
        }

        private void FrmXML_Respuestas_FormClosing(object sender, FormClosingEventArgs e)
        {
            usage.End("Contratación", "FrmXML_Respuestas" ,"N/A");
        }

        private void label1_Click(object sender, EventArgs e)
        {
         
        }

        //rivate void button1_Click(object sender, EventArgs e)  // boton C102
        private void btn_c102_Click(object sender, EventArgs e)
        {

            // MessageBox.Show("Este botón o proceso aún no está en funcionamiento.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
             EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml;
  ;

            foreach (DataGridViewRow row in dgv.SelectedRows)
            {

                DataGridViewCell id = (DataGridViewCell)row.Cells[0];

                xml = cont_xml.CargaDatosC1(Convert.ToInt32(id.Value.ToString()));
                //if (xml == null || xml.Cabecera == null)
                //{
                //    MessageBox.Show("ERROR: No se han recuperado datos del mensaje C101.");
                //    return;
                //}
                //else
                //{
                //    MessageBox.Show(
                //        $"XML cargado correctamente.\n" +
                //        $"Código Solicitud: {xml.Cabecera.CodigoDeSolicitud}\n" +
                //        $"CUPS: {xml.Cabecera.CUPS}");
                //}



                FrmXML_Respuestas_Datos_C102 f =
                    new FrmXML_Respuestas_Datos_C102(xml, cnmc, cont_xml);

                f.ShowDialog();


                MessageBox.Show(@"Mensajes creados en c:\Temp",
               "Creación Paso C102",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information);
            }
        }

        private void btn_C105_Click(object sender, EventArgs e)
        {
            // MessageBox.Show("Este botón o proceso aún no está en funcionamiento.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
            EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml;

            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                DataGridViewCell id = (DataGridViewCell)row.Cells[0];

                xml = cont_xml.CargaDatosC1(Convert.ToInt32(id.Value.ToString()));

                FrmXML_Respuestas_Datos_C105 f =
                    new FrmXML_Respuestas_Datos_C105(xml, cnmc, cont_xml);

                f.ShowDialog();

                MessageBox.Show(@"Mensajes creados en c:\Temp",
                   "Creación Paso C105",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
            }



        }

        private void button5_Click(object sender, EventArgs e) //  boton  m102
        
        {
            EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml;
               
            

            foreach (DataGridViewRow row in dgv.SelectedRows)
            {

                DataGridViewCell id = (DataGridViewCell)row.Cells[0];

                xml = cont_xml.CargaDatosM1(Convert.ToInt32(id.Value.ToString()));
                //if (xml == null || xml.Cabecera == null)
                //{
                //    MessageBox.Show("ERROR: No se han recuperado datos del mensaje C101.");
                //    return;
                //}
                //else
                //{
                //    MessageBox.Show(
                //        $"XML cargado correctamente.\n" +
                //        $"Código Solicitud: {xml.Cabecera.CodigoDeSolicitud}\n" +
                //        $"CUPS: {xml.Cabecera.CUPS}");
                //}



                // FrmXML_Respuestas_Datos_C102 f =
                //     new FrmXML_Respuestas_Datos_C102(xml, cnmc, cont_xml);
                FrmXML_Respuestas_Datos_M102 f =
                    new FrmXML_Respuestas_Datos_M102(xml, cnmc, cont_xml);
                 f.ShowDialog();


                MessageBox.Show(@"Mensajes creados en c:\Temp",
               "Creación Paso M102",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information);
            }
        }

        private void btn_b102_Click(object sender, EventArgs e)
        {

        }

        private void btn_b105_Click(object sender, EventArgs e)
        {

        }

        private void btn_c202_Click(object sender, EventArgs e)
        {

        }

        private void btn_c205_Click(object sender, EventArgs e)
        {

        }
    }
}
