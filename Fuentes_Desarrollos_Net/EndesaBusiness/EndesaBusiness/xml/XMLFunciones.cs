using EndesaBusiness.contratacion.eexxi;
using EndesaBusiness.servidores;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.facturacion;
using iTextSharp.text.pdf.parser;

//using Microsoft.Graph;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace EndesaBusiness.xml
{


    public class XMLFunciones
    {



        Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_t101_xml;
        Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_t105_xml;
        Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_xml;
        EndesaBusiness.utilidades.Param p;
        EndesaBusiness.contratacion.eexxi.EEXXI xxi;
        EndesaBusiness.global.Municipios municipio;
        EndesaBusiness.global.Provincias provincias;


        public XMLFunciones()
        {

            dic_t101_xml = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();
            dic_t105_xml = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();
            dic_xml = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();

            EndesaEntity.contratacion.xxi.XML_Medidas med;
            municipio = new global.Municipios();
            provincias = new global.Provincias("eexxi_param_provincias",EndesaBusiness.servidores.MySQLDB.Esquemas.CON);

            EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();

            p = new EndesaBusiness.utilidades.Param("eexxi_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);



            #region Ejemplo datos XML
            //xml.codigoREEEmpresaEmisora = "0023";
            //xml.codigoREEEmpresaDestino = "0636";
            //xml.codigoDelProceso = "T1";
            //xml.codigoDePaso = "05";
            //xml.codigoDeSolicitud = "201800002812";
            //xml.secuencialDeSolicitud = "01";
            //xml.fechaSolicitud = new DateTime(2020, 09, 01, 14, 13, 26);
            //xml.cups = "ES0031101562016001JS0F";
            //xml.cnae = 1013;
            //xml.potenciaExtension = 23098;
            //xml.potenciaDeAcceso = 220000;
            //xml.indicativoDeInterrumpibilidad = "N";
            //xml.pais = "ES";
            //xml.provincia = "06";
            //xml.poblacion = "061160001";
            //xml.descripcionPoblacion = "SALVALEON";
            //xml.tipoVia = "CR";
            //xml.codPostal = "06174";
            //xml.calle = "ZAFRA";
            //xml.numeroFinca = "SN";
            //xml.piso = "LOC";
            //xml.tipoAclaradorFinca = "KM";
            //xml.aclaradorFinca = "3";

            //xml.tipoIdentificador = "NI";
            //xml.identificador = "F06021224";
            //xml.tipoPersona = "J";
            //xml.razonSocial = "MONTEPORRINO S C L";
            //xml.prefijoPais = "34";
            //xml.numero = "924752561";
            //xml.indicadorTipoDireccion = "F";
            //xml.indicativoDeDireccionExterna = "S";
            //xml.linea1DeLaDireccionExterna = "CR ZAFRA LOC KM 3 SALVALE";
            //xml.linea2DeLaDireccionExterna = "06174 SALVALEON";
            //xml.linea3DeLaDireccionExterna = "Badajoz-España";
            //xml.linea4DeLaDireccionExterna = "ON";

            //xml.idioma = "ES";
            //xml.fechaActivacion = new DateTime(2019, 01, 06);
            //xml.codContrato = "500002383556";
            //xml.tipoAutoconsumo = "00";
            //xml.tipoContratoATR = "01";
            //xml.tarifaATR = "011";
            //xml.periodicidadFacturacion = "01";
            //xml.tipodeTelegestion = "03";
            //xml.potenciaPeriodo[1] = 220000;
            //xml.potenciaPeriodo[2] = 220000;
            //xml.potenciaPeriodo[3] = 220000;
            //xml.marcaMedidaConPerdidas = "N";
            //xml.tensionDelSuministro = 24;

            //EndesaEntity.contratacion.xxi.XML_PuntoDeMedida pm = new EndesaEntity.contratacion.xxi.XML_PuntoDeMedida();
            //pm.codPM = "ES0031101562016001JS1P";
            //pm.tipoMovimiento = "B";
            //pm.tipoPM = "03";
            //pm.modoLectura = "4";
            //pm.funcion = "P";
            //pm.telefonoTelemedida = "000000000";
            //pm.tensionPM = "24";
            //pm.fechaVigor = new DateTime(2018, 01, 01);
            //pm.fechaAlta = new DateTime(2008, 06, 05);
            //pm.fechaBaja = new DateTime(2019, 01, 05);

            //EndesaEntity.contratacion.xxi.XML_Aparatos aparato = new EndesaEntity.contratacion.xxi.XML_Aparatos();
            //aparato.tipoAparato = "TI";
            //aparato.marcaAparato = "003";
            //aparato.modeloMarca = "J64R87";
            //aparato.tipoMovimiento = "CX";
            //aparato.tipoEquipoMedida = "L03";
            //aparato.tipoPropiedadAparato = "2";
            //aparato.tipoDHEdM = "1";
            //aparato.modoMedidaPotencia = "1";
            //aparato.codPrecinto = "20629";
            //aparato.periodoFabricacion = "1990";
            //aparato.numeroSerie = "005266226";
            //aparato.funcionAparato = "M";
            //aparato.numIntegradores = "0";
            //aparato.constanteEnergia = "0.000";
            //aparato.constanteMaximetro = "0.000";
            //aparato.ruedasEnteras = "0";
            //aparato.ruedasDecimales = "0";
            //pm.lista_aparatos.Add(aparato);

            //aparato = new EndesaEntity.contratacion.xxi.XML_Aparatos();
            //aparato.tipoAparato = "TI";
            //aparato.marcaAparato = "003";
            //aparato.modeloMarca = "J64R87";
            //aparato.tipoMovimiento = "CX";
            //aparato.tipoEquipoMedida = "L03";
            //aparato.tipoPropiedadAparato = "2";
            //aparato.tipoDHEdM = "1";
            //aparato.modoMedidaPotencia = "1";
            //aparato.codPrecinto = "20629";
            //aparato.periodoFabricacion = "1990";
            //aparato.numeroSerie = "005266226";
            //aparato.funcionAparato = "M";
            //aparato.numIntegradores = "0";
            //aparato.constanteEnergia = "0.000";
            //aparato.constanteMaximetro = "0.000";
            //aparato.ruedasEnteras = "0";
            //aparato.ruedasDecimales = "0";
            //pm.lista_aparatos.Add(aparato);



            //med = new EndesaEntity.contratacion.xxi.XML_Medidas();
            //med.tipoDHEdM = "3";
            //med.periodo = "32";
            //med.magnitudMedida = "AE";
            //med.procedencia = "20";
            //med.ultimaLecturaFirme = "4797630.00";
            //med.fechaLecturaFirme = new DateTime(2018, 11, 30);
            //aparato.listaMedidas.Add(med);


            //pm.lista_aparatos.Add(aparato);
            //xml.lista_PM.Add(pm);

            #endregion


            //CreaXML(xml);
        }

        public void CargaDatosExcel_XML_Platillas(string fichero_excel, string plantilla_xml, string directorio_salida)
        {
            int id = 0;
            int c = 1;
            int f = 1;
            bool firstOnly = true;
            List<string> lista_cabecera = new List<string>();
            Dictionary<string, List<EndesaEntity.xml.Etiqueta_valor_XML>> dic =
                new Dictionary<string, List<EndesaEntity.xml.Etiqueta_valor_XML>>();
            string clave = "";
            string posicion_cabecera = "";

            try
            {

                posicion_cabecera = plantilla_xml.Substring(plantilla_xml.LastIndexOf('_') + 1,
                    plantilla_xml.Length - (plantilla_xml.LastIndexOf('.') + 1));

                FileStream fs = new FileStream(fichero_excel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fs);
                var workSheet = excelPackage.Workbook.Worksheets.First();

                for (int i = 0; i < 1000000; i++)
                {
                    c = 1;

                    if (firstOnly)
                    {
                        for (int w = 1; w < 1000; w++)
                        {
                            if (workSheet.Cells[f, w].Value == null)
                                break;

                            lista_cabecera.Add(workSheet.Cells[f, w].Value.ToString());
                        }

                        firstOnly = false;
                    }

                    f++;

                    if (workSheet.Cells[f, 1].Value == null)
                        break;
                    else
                    {
                        clave = "";
                                                                     

                        foreach (string p in lista_cabecera)
                        {
                            for (int z = 0; z < posicion_cabecera.Length; z++)
                                if (c.ToString() == Convert.ToString(posicion_cabecera[z]))
                                    clave += workSheet.Cells[f, c].Value.ToString();

                                //clave += workSheet.Cells[f, c].Value.ToString();
                            c++;
                        }

                        c = 1;


                        List<EndesaEntity.xml.Etiqueta_valor_XML> lista =
                            new List<EndesaEntity.xml.Etiqueta_valor_XML>();
                        foreach (string p in lista_cabecera)
                        {
                            EndesaEntity.xml.Etiqueta_valor_XML tupla =
                                new EndesaEntity.xml.Etiqueta_valor_XML();
                            tupla.etiqueta = p;
                            tupla.valor = workSheet.Cells[f, c].Value.ToString();
                            c++;
                            lista.Add(tupla);
                        }

                        dic.Add(clave, lista);


                    }
                }


                fs = null;
                excelPackage = null;

                string etiqueta = "";
                string resultado = "";

                // Por cada datos del Dic creamos un XML
                foreach(KeyValuePair<string, List<EndesaEntity.xml.Etiqueta_valor_XML>> p in dic)
                {
                    string str = File.ReadAllText(plantilla_xml);

                    foreach (EndesaEntity.xml.Etiqueta_valor_XML pp in p.Value)
                    {
                        etiqueta = "<" + pp.etiqueta + ">" + "</" + pp.etiqueta + ">";
                        resultado = "<" + pp.etiqueta + ">" + pp.valor + "</" + pp.etiqueta + ">";
                        str = str.Replace(etiqueta, resultado);                            
                    }

                    File.WriteAllText(directorio_salida + "\\" + p.Key + ".xml", str);
                }


                MessageBox.Show("Proceso finalizado",
               "Generar XML desde plantilla",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information);


            }
            catch (Exception e)
            {

            }
        }



        public void CreaRespuestas()
        {

        }



        //public void CargaXML()
        //{

        //    double percent = 0;
        //    int f_altas = 0;            
        //    int total_archivos = 0;
        //    DialogResult result = DialogResult.Yes;


        //    List<EndesaEntity.contratacion.xxi.XML_Datos> lista = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
        //    EndesaBusiness.cups.TarifaATR tarifaATR = new EndesaBusiness.cups.TarifaATR();
        //    XML_SolicitudesCodigos sol_codigos = new XML_SolicitudesCodigos();

        //    OpenFileDialog d = new OpenFileDialog();
        //    // d.Title = p.GetValue("mensaje_ventana_xml", DateTime.Now, DateTime.Now);
        //    d.Title = "Seleccione ruta de entrada de archivos XML";
        //    //d.Filter = "zip files|*.zip";
        //    d.Filter = "XML files|*.xml";
        //    d.Multiselect = true;
        //    List<string> files = new List<string>();
        //    List<string> files2 = new List<string>();

        //    if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)            
        //    {
                

        //        if (result == DialogResult.Yes)
        //        {
        //            List<EndesaEntity.contratacion.Solicitudes_Codigos_Tabla> lista_altas =
        //                sol_codigos.dic.Where(z => z.Value.descripcion == "ALTA").Select(z => z.Value).ToList();

        //            total_archivos = d.FileNames.Count();
        //            f_altas = 0;

        //            Cursor.Current = Cursors.WaitCursor;
        //            forms.FrmProgressBar pb = new forms.FrmProgressBar();

        //            foreach (string fileName in d.FileNames)
        //            {

        //                files.Add(fileName);

        //            }
                       

        //            for (int z = 0; z < lista_altas.Count; z++)
        //            {
        //                files2.Clear();

        //                for(int j = 0; j < files.Count; j++)
        //                {
        //                    if (files[j].Contains(lista_altas[z].codigoproceso + "_" + lista_altas[z].codigopaso))
        //                        files2.Add(files[j]);
        //                }                           

        //                pb.Text = "Carga archivos XML";
        //                pb.Show();
        //                pb.progressBar.Step = 1;
        //                pb.progressBar.Maximum = files.Count();

        //                for (int i = 0; i < files2.Count(); i++)
        //                {

        //                    f_altas++;
        //                    FileInfo file = new FileInfo(files2[i]);

        //                    percent = (i / Convert.ToDouble(files2.Count())) * 100;
        //                    pb.progressBar.Increment(1);
        //                    pb.txtDescripcion.Text = "Importando ALTAS " + lista_altas[z].codigoproceso + " "
        //                        + i.ToString("#.###") + " / " + files.Count().ToString("#.###") + " archivos --> " + file.Name;
        //                    pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
        //                    pb.Refresh();


        //                    if (file.Name.Contains(lista_altas[z].codigoproceso + "_" + lista_altas[z].codigopaso))
        //                    {
        //                        EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();
        //                        xml = TrataXML(file.FullName);

        //                        if (tarifaATR.EsTarifaEEXXI(xml.tarifaATR) && (xml.potenciaPeriodo[3] > 50000 || xml.potenciaPeriodo[6] > 50000))
        //                            if (z == 0)
        //                                dic_t101_xml.Add(xml.codigoDeSolicitud, xml);
        //                            else
        //                                dic_t105_xml.Add(xml.codigoDeSolicitud, xml);

        //                    }


        //                }
                            
        //            }

        //            pb.Close();
                   

        //            if (!Directory.Exists(@"C:\Temp\t305_formato_antiguo\"))
        //                Directory.CreateDirectory(@"C:\Temp\t305_formato_antiguo\");


        //            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.XML_Datos> p in dic_t101_xml)
        //            {
        //                EndesaEntity.contratacion.xxi.XML_Datos oXML;
        //                if (dic_t105_xml.TryGetValue(p.Key, out oXML))
        //                {
        //                    FileInfo ficheroSalida = new FileInfo(@"C:\Temp\t305_formato_antiguo\" + oXML.fichero);
        //                    CreaXML_T105_Clasic(ficheroSalida, p.Value, oXML);

        //                }
                        
        //            }


        //            GuardadoBBDD_T101(dic_t101_xml);

        //            MessageBox.Show("La transformación ha concluido correctamente."
        //            + System.Environment.NewLine
        //            + System.Environment.NewLine 
        //            + @"Los resultados se han guardado en C:\Temp\t305_formato_antiguo\",
        //           "Importación ficheros XML",
        //           MessageBoxButtons.OK,
        //           MessageBoxIcon.Information);





        //        }

                
        //    }
        //}

        public void CargaXML_A302_A305()
        {

            double percent = 0;
            int f_altas = 0;
            int total_archivos = 0;
            DialogResult result = DialogResult.Yes;
            bool error = true;

            EndesaEntity.contratacion.xxi.XML_Datos o;

            List<EndesaEntity.contratacion.xxi.XML_Datos> lista = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            EndesaBusiness.cups.TarifaATR tarifaATR = new EndesaBusiness.cups.TarifaATR();
            XML_SolicitudesCodigos sol_codigos = new XML_SolicitudesCodigos();

            OpenFileDialog d = new OpenFileDialog();
            
            d.Title = "Seleccione los archivos XML de los pasos A301, A302 y A305";            
            d.Filter = "XML files|*.xml";
            d.Multiselect = true;
            List<string> files = new List<string>();
            List<string> files2 = new List<string>();

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {


                if (result == DialogResult.Yes)
                {
                    List<EndesaEntity.contratacion.Solicitudes_Codigos_Tabla> lista_altas =
                        sol_codigos.dic.Where(z => z.Value.descripcion == "ALTA_ESPECIAL").Select(z => z.Value).ToList();

                    total_archivos = d.FileNames.Count();
                    f_altas = 0;

                    Cursor.Current = Cursors.WaitCursor;
                    forms.FrmProgressBar pb = new forms.FrmProgressBar();

                    foreach (string fileName in d.FileNames)
                    {
                        files.Add(fileName);
                    }




                    for (int z = 0; z < lista_altas.Count; z++)
                    {
                        files2.Clear();

                        for (int j = 0; j < files.Count; j++)
                        {
                            //if (files[j].Contains(lista_altas[z].codigoproceso + "_" + lista_altas[z].codigopaso))
                                files2.Add(files[j]);
                        }

                        pb.Text = "Carga archivos XML";
                        pb.Show();
                        pb.progressBar.Step = 1;
                        pb.progressBar.Maximum = files.Count();

                        for (int i = 0; i < files2.Count(); i++)
                        {

                            f_altas++;
                            FileInfo file = new FileInfo(files2[i]);

                            percent = (i / Convert.ToDouble(files2.Count())) * 100;
                            pb.progressBar.Increment(1);
                            pb.txtDescripcion.Text = "Importando ALTAS " + lista_altas[z].codigoproceso + " "
                                + i.ToString("#.###") + " / " + files.Count().ToString("#.###") + " archivos --> " + file.Name;
                            pb.lbl_percent.Text = string.Format("{0}", percent.ToString("##0") + " %");
                            pb.Refresh();


                            //if (file.Name.Contains(lista_altas[z].codigoproceso + "_" + lista_altas[z].codigopaso))
                           // {
                            EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();
                            xml = TrataXML(file.FullName);

                            if (tarifaATR.EsTarifaEEXXI(xml.tarifaATR) && (xml.potenciaPeriodo[3] > 50000 || xml.potenciaPeriodo[6] > 50000))
                            {
                                if (z == 0)
                                {
                                    if(!dic_t101_xml.TryGetValue(xml.codigoDeSolicitud, out o))
                                        dic_t101_xml.Add(xml.codigoDeSolicitud, xml);
                                }
                                else
                                {
                                    if (!dic_t105_xml.TryGetValue(xml.codigoDeSolicitud, out o))
                                        dic_t105_xml.Add(xml.codigoDeSolicitud, xml);
                                }                                   

                            }




                            //  }

                        }

                    }

                    pb.Close();


                    if (!Directory.Exists(@"C:\Temp\XML\"))
                        Directory.CreateDirectory(@"C:\Temp\XML\");


                    foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.XML_Datos> pp in dic_t101_xml)
                    {
                        EndesaEntity.contratacion.xxi.XML_Datos oXML;
                        if (dic_t105_xml.TryGetValue(pp.Key, out oXML))
                        {
                            FileInfo ficheroSalida = new FileInfo(@"C:\Temp\XML\" + oXML.fichero.Replace("A3_05","T1_05"));
                            error = CreaXML_t105_A302_A305_v2(ficheroSalida, pp.Value, oXML);

                            ficheroSalida = new FileInfo(@"C:\Temp\XML\" + oXML.fichero.Replace("A3_05", "T1_01"));
                            error = CreaXML_t101_A302_A305_v2(ficheroSalida, pp.Value, oXML);

                            //XmlReaderSettings settings = new XmlReaderSettings();
                            //settings.ValidationType = ValidationType.Schema;
                            //settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
                            //settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessSchemaLocation;
                            //settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                            //settings.ValidationEventHandler += new ValidationEventHandler(ValidationCallBack);

                            //XmlReader rd = XmlReader.Create(ficheroSalida.FullName, settings);

                            //XmlSchemaSet schema = new XmlSchemaSet();
                            //schema.Add("", System.Environment.CurrentDirectory 
                            //    + p.GetValue("xsd_t105"));

                            //XDocument doc = XDocument.Load(rd);
                            //doc.Validate(schema, ValidationEventHandler);
                        }
                        else
                        {
                            MessageBox.Show("No se ha podido encontrar el T101 para el CUPS "
                                + pp.Value.cups + "."
                               + System.Environment.NewLine
                               + System.Environment.NewLine
                               + "No se generará el fichero XML.",
                              "Generación T105",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
                            error = true;
                        }

                    }

                    // GuardadoBBDD_T101(dic_t101_xml);
                    if (!error)
                    {
                        MessageBox.Show("La transformación ha concluido correctamente."
                            + System.Environment.NewLine
                            + System.Environment.NewLine
                            + @"Los resultados se han guardado en C:\Temp\XML\",
                           "Importación ficheros XML",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Information);
                    }
                    





                }


            }
        }

        private EndesaEntity.contratacion.xxi.XML_Datos TrataXML(string fileName)
        {
            string cod_ini = "";
            string cod_fin = "";
            string valor = "";
            bool dentroCliente = false;
            int numCodPM = 0;
            int numTipoAparato = 0;
            string dentroDe = "";

            int potencia = 0;

            EndesaEntity.contratacion.xxi.XML_Datos xml = new EndesaEntity.contratacion.xxi.XML_Datos();

            EndesaEntity.contratacion.xxi.XML_PuntoDeMedida pm
                = new EndesaEntity.contratacion.xxi.XML_PuntoDeMedida();

            EndesaEntity.contratacion.xxi.XML_Aparatos aparato
                = new EndesaEntity.contratacion.xxi.XML_Aparatos();

            FileInfo file = new FileInfo(fileName);

            xml.fichero = file.Name;

            XmlTextReader r = new XmlTextReader(fileName);
            while (r.Read())
            {

                switch (r.NodeType)
                {

                    case XmlNodeType.Element: // The node is an element.

                        cod_ini = r.Name;

                        if (!dentroCliente)
                            dentroCliente = (cod_ini == "Direccion");

                        if (cod_ini == "Potencia")
                            potencia++;

                        switch (cod_ini)
                        {
                            
                            case "PuntoDeMedida":
                                pm = new EndesaEntity.contratacion.xxi.XML_PuntoDeMedida();
                                dentroDe = cod_ini;
                                break;
                            case "Aparato":
                                aparato = new EndesaEntity.contratacion.xxi.XML_Aparatos();
                                dentroDe = cod_ini;
                                break;

                            case "Cliente":
                                dentroDe = cod_ini;
                                break;
                            case "DireccionPS":
                                dentroDe = cod_ini;
                                break;
                            case "Contacto":
                                dentroDe = cod_ini;
                                break;

                        }
                        
                        break;
                    case XmlNodeType.Text: //Display the text in each element.
                        valor = EndesaBusiness.utilidades.FuncionesTexto.ArreglaAcentos(r.Value);
                        break;
                    case XmlNodeType.EndElement: //Display the end of the element.
                        cod_fin = r.Name;

                        switch (cod_fin)
                        {
                            case "PuntoDeMedida":
                                xml.lista_PM.Add(pm);
                                break;
                            case "Aparato":
                                pm.lista_aparatos.Add(aparato);
                                break;
                        }

                        break;


                }

                #region XML

                if (cod_ini == cod_fin)
                    switch (cod_ini)
                    {
                        case "NombreDePila":
                            xml.razonSocial = valor;
                            break;
                        case "PrimerApellido":
                            xml.razonSocial += " " + valor;
                            break;
                        case "CodigoREEEmpresaEmisora":
                            xml.codigoREEEmpresaEmisora = valor;
                            break;
                        case "CodigoREEEmpresaDestino":
                            xml.codigoREEEmpresaDestino = valor;
                            break;
                        case "CodigoDelProceso":
                            xml.codigoDelProceso = valor;
                            break;
                        case "CodigoDePaso":
                            xml.codigoDePaso = valor;
                            break;
                        case "CodigoDeSolicitud":
                            xml.codigoDeSolicitud = valor;
                            break;
                        case "SecuencialDeSolicitud":
                            xml.secuencialDeSolicitud = valor;
                            break;
                        case "FechaSolicitud":
                            xml.fechaSolicitud = Convert.ToDateTime(valor.Substring(0, 10) + " " + valor.Substring(11, 8));
                            break;
                        case "CUPS":
                            xml.cups = valor;
                            break;
                        case "CNAE":
                            xml.cnae = valor;
                            break;
                        case "PotenciaExtension":
                            xml.potenciaExtension = Convert.ToDouble(valor);
                            break;
                        case "PotenciaDeAcceso":
                            xml.potenciaDeAcceso = Convert.ToDouble(valor);
                            break;
                        case "PotenciaInstAT":
                            xml.potenciaInstAT = Convert.ToDouble(valor);
                            break;
                        case "IndicativoDeInterrumpibilidad":
                            xml.indicativoDeInterrumpibilidad = valor;
                            break;
                        case "Pais":
                            if (xml.pais != null)
                                xml.paisCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_Pais = valor;
                            else
                                xml.pais = valor;
                            break;
                        case "Provincia":
                            if (xml.provincia != null)
                                xml.provinciaCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_Provincia = valor;
                            else
                                xml.provincia = valor;
                            break;
                        case "Municipio":
                            if (xml.municipio != null)
                                xml.municipioCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_Municipio = valor;
                            else
                                xml.municipio = valor;
                            break;
                        case "Poblacion":
                            if (xml.poblacion != null)
                                xml.poblacionCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_Poblacion = valor;
                            else
                                xml.poblacion = valor;
                            break;
                        case "FechaPrevistaAccion":                            
                                xml.fechaPrevistaAccion = Convert.ToDateTime(valor);
                            break;
                        case "MotivoTraspaso":
                            xml.motivoTraspaso = valor;
                            break;
                        case "IndEsencial":
                            xml.indEsencial = valor;
                            break;
                        case "DescripcionPoblacion":
                            if (xml.descripcionPoblacion != null)
                                xml.descripcionPoblacionCliente = valor;
                            else
                                xml.descripcionPoblacion = valor;
                            break;
                        case "TipoVia":
                            if (dentroDe == "Cliente")
                                xml.tipoViaCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_TipoVia = valor;
                            else
                                xml.tipoVia = valor;
                            break;
                        case "CodPostal":
                            if (xml.codPostal != null)
                                xml.codPostalCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_CodPostal = valor;
                            else
                                xml.codPostal = valor;
                            break;
                        case "Calle":
                            if (dentroDe == "Cliente")
                                xml.calleCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_Calle = valor;
                            else
                                xml.calle = valor;
                            break;
                        case "NumeroFinca":
                            if (dentroDe == "Cliente")
                                xml.numeroFincaCliente = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_NumeroFinca = valor;
                            else
                                xml.numeroFinca = valor;
                            break;
                        case "AclaradorFinca":
                            if (dentroDe == "Cliente")
                                xml.aclaradorFinca = valor;
                            if (dentroDe == "DireccionPS")
                                xml.direccionPS_AclaradorFinca = valor;
                            break;
                        case "TipoIdentificador":                            
                            xml.tipoIdentificador = valor;
                            break;
                        case "Identificador":
                            xml.identificador = valor;
                            break;
                        case "TipoPersona":
                            xml.tipoPersona = valor;
                            break;
                        case "RazonSocial":
                            xml.razonSocial = valor;
                            break;
                        case "PerdonaDeContacto":
                            xml.contacto_PersonaDeContacto = valor;
                            break;
                        case "PrefijoPais":
                            if (dentroDe == "Contacto")
                                xml.contacto_PrefijoPais = valor;
                            else
                                xml.prefijoPais = valor;
                            break;
                        case "Numero":
                            if (dentroDe == "Contacto")
                                xml.contacto_Numero = valor;
                            else 
                                xml.numero = valor;
                            break;
                        case "CorreoElectronico":
                            xml.correoElectronico = valor;
                            break;
                        case "IndicadorTipoDireccion":
                            xml.indicadorTipoDireccion = valor;                            
                            break;
                        case "IndicativoDeDireccionExterna":
                            xml.indicativoDeDireccionExterna = valor;
                            break;
                        case "Linea1DeLaDireccionExterna":
                            xml.linea1DeLaDireccionExterna = valor;
                            break;
                        case "Linea2DeLaDireccionExterna":
                            xml.linea2DeLaDireccionExterna = valor;
                            break;
                        case "Linea3DeLaDireccionExterna":
                            xml.linea3DeLaDireccionExterna = valor;
                            break;
                        case "Linea4DeLaDireccionExterna":
                            xml.linea4DeLaDireccionExterna = valor;
                            break;
                        case "Idioma":
                            xml.idioma = valor;
                            break;
                        case "Fecha":
                            xml.fechaActivacion = Convert.ToDateTime(valor);
                            break;
                        case "FechaActivacion":
                            xml.fechaActivacion = Convert.ToDateTime(valor);
                            break;
                        case "CodContrato":
                            xml.codContrato = valor;
                            break;
                        case "TipoAutoconsumo":
                            xml.tipoAutoconsumo = valor;
                            break;
                        case "TipoContratoATR":
                            xml.tipoContratoATR = valor;
                            break;
                        case "TarifaATR":
                            xml.tarifaATR = valor;
                            break;
                        case "PeriodicidadFacturacion":
                            xml.periodicidadFacturacion = valor;
                            break;
                        case "TipodeTelegestion":
                            xml.tipodeTelegestion = valor;
                            break;
                        case "Potencia":
                            xml.potenciaPeriodo[potencia] = Convert.ToDouble(valor);
                            break;
                        case "MarcaMedidaConPerdidas":
                            xml.marcaMedidaConPerdidas = valor;
                            break;
                        case "TensionDelSuministro":
                            xml.tensionDelSuministro = Convert.ToInt32(valor);
                            break;
                        case "VAsTrafo":
                            xml.vasTrafo = Convert.ToInt32(valor);
                            break;
                        case "CodPM":
                            numCodPM++;
                            
                            if(numCodPM > 1)
                                pm = new EndesaEntity.contratacion.xxi.XML_PuntoDeMedida();

                            pm.codPM = valor;
                            break;
                        case "TipoMovimiento":
                            if (dentroDe == "PuntoDeMedida")
                                pm.tipoMovimiento = valor;
                            else
                                aparato.tipoMovimiento = valor;
                            break;
                        case "TipoPM":
                            pm.tipoPM = valor;
                            break;
                        case "ModoLectura":
                            pm.modoLectura = valor;
                            break;
                        case "Funcion":
                            pm.funcion = valor;
                            break;
                        case "TelefonoTelemedida":
                            pm.telefonoTelemedida = valor;
                            break;                        
                        case "TensionPM":
                            pm.tensionPM = valor;
                            break;
                        case "FechaVigor":
                            pm.fechaVigor = Convert.ToDateTime(valor);
                            break;
                        case "FechaAlta":
                            pm.fechaAlta = Convert.ToDateTime(valor);
                            break;
                        case "FechaBaja":
                            pm.fechaBaja = Convert.ToDateTime(valor);
                            break;
                        case "TipoAparato":                           

                            aparato.tipoAparato = valor;
                            break;

                        case "MarcaAparato":
                            aparato.marcaAparato = valor;
                            break;
                        case "ModeloMarca":
                            aparato.modeloMarca = valor;
                            break;                       
                        case "TipoEquipoMedida":
                            aparato.tipoEquipoMedida = valor;
                            break;
                        case "TipoPropiedadAparato":
                            aparato.tipoPropiedadAparato = valor;
                            break;
                        case "TipoDHEdM":
                            aparato.tipoDHEdM = valor;
                            break;
                        case "ModoMedidaPotencia":
                            aparato.modoMedidaPotencia = valor;
                            break;
                        case "CodPrecinto":
                            aparato.codPrecinto = valor;
                            break;
                        case "PeriodoFabricacion":
                            aparato.periodoFabricacion = valor;
                            break;
                        case "NumeroSerie":
                            aparato.numeroSerie = valor;
                            break;
                        case "FuncionAparato":
                            aparato.funcionAparato = valor;
                            break;
                        case "NumIntegradores":
                            aparato.numIntegradores = valor;
                            break;
                        case "ConstanteEnergia":
                            aparato.constanteEnergia = valor;
                            break;
                        case "ConstanteMaximetro":
                            aparato.constanteMaximetro = valor;
                            break;
                        case "RuedasEnteras":
                            aparato.ruedasEnteras = valor;
                            break;
                        case "RuedasDecimales":
                            aparato.ruedasDecimales = valor;
                            break;


                    }


                #endregion

            }

            if(xml.indicadorTipoDireccion == "S")
            {
                xml.calleCliente = xml.direccionPS_Calle;                
                xml.numeroFincaCliente = xml.direccionPS_NumeroFinca;
                xml.codPostalCliente = xml.direccionPS_CodPostal;
                xml.municipioCliente = xml.direccionPS_Municipio;
                xml.provinciaCliente = xml.direccionPS_Provincia;
                xml.descripcionPoblacionCliente = xml.direccionPS_Poblacion;
                xml.paisCliente = xml.direccionPS_Pais;
                xml.tipoViaCliente = xml.direccionPS_TipoVia;
            }

            return xml;


        }

        //public void CreaXML_T105_Clasic(FileInfo file, EndesaEntity.contratacion.xxi.XML_Datos t101, EndesaEntity.contratacion.xxi.XML_Datos t105)
        //{

        //    string url = "http://xmlns.endesa.com/wsdl/FormatoCUR/";

        //    XmlTextWriter writer = new XmlTextWriter(file.FullName, Encoding.UTF8);


        //    writer.WriteStartDocument();
        //    writer.WriteStartElement("VueltasaCUR", url);
        //    {
        //        writer.WriteStartElement("Cabecera");
        //        {
        //            writer.WriteElementString("CodigoREEEmpresaEmisora", t101.codigoREEEmpresaEmisora);
        //            writer.WriteElementString("CodigoREEEmpresaDestino", t101.codigoREEEmpresaDestino);
        //            writer.WriteElementString("CodigoDelProceso", t105.codigoDelProceso);
        //            writer.WriteElementString("CodigoDePaso", t105.codigoDePaso);
        //            writer.WriteElementString("CodigoDeSolicitud", t105.codigoDeSolicitud);
        //            writer.WriteElementString("SecuencialDeSolicitud", t105.secuencialDeSolicitud);
        //            writer.WriteElementString("FechaSolicitud", t101.fechaSolicitud.ToString("yyyy-MM-ddTHH:mm:ss"));
        //            writer.WriteElementString("CUPS", t101.cups);
        //        }
        //        writer.WriteEndElement(); // Cabecera

        //        writer.WriteStartElement("RegistrodePuntodeSuministro");
        //        {
        //            writer.WriteStartElement("Suministro");
        //            {
        //                writer.WriteElementString("CNAE", t101.cnae.ToString());
        //                writer.WriteElementString("PotenciaExtension", t105.potenciaExtension.ToString());
        //                writer.WriteElementString("PotenciaDeAcceso", t105.potenciaDeAcceso.ToString());
        //                writer.WriteElementString("PotenciaInstAT", t105.potenciaInstAT.ToString());
        //                writer.WriteElementString("IndicativoDeInterrumpibilidad", t105.indicativoDeInterrumpibilidad);

        //                writer.WriteStartElement("DireccionPS");
        //                {
        //                    writer.WriteElementString("Pais", t101.direccionPS_Pais);
        //                    writer.WriteElementString("Provincia", provincias.DesProvincia(t101.direccionPS_Provincia));
        //                    writer.WriteElementString("Municipio", municipio.DesMunicipio(t101.direccionPS_Municipio));
        //                    writer.WriteElementString("Poblacion", t101.poblacion);
        //                    writer.WriteElementString("DescripcionPoblacion", t101.descripcionPoblacion);
        //                    writer.WriteElementString("TipoVia", t101.tipoVia);
        //                    writer.WriteElementString("CodPostal", t101.direccionPS_CodPostal);
        //                    writer.WriteElementString("Calle", t101.direccionPS_Calle);
        //                    writer.WriteElementString("NumeroFinca", t101.direccionPS_NumeroFinca);
        //                    writer.WriteElementString("Piso", "");
        //                    writer.WriteElementString("TipoAclaradorFinca", "");
        //                    writer.WriteElementString("AclaradorFinca", t101.direccionPS_AclaradorFinca);
        //                }
        //                writer.WriteEndElement(); // DireccionPS

        //            }
        //            writer.WriteEndElement(); // Suministro

        //            writer.WriteStartElement("Cliente");
        //            {
        //                writer.WriteStartElement("IdCliente");
        //                {
        //                    writer.WriteElementString("TipoIdentificador", t101.tipoIdentificador);
        //                    writer.WriteElementString("Identificador", t101.identificador);
        //                    writer.WriteElementString("TipoPersona", t101.tipoPersona);
        //                }
        //                writer.WriteEndElement(); // IdCliente

        //                writer.WriteStartElement("Nombre");
        //                {
        //                    writer.WriteElementString("RazonSocial", t101.razonSocial);
        //                }
        //                writer.WriteEndElement(); // Nombre

        //                writer.WriteStartElement("Telefono");
        //                {
        //                    writer.WriteElementString("PrefijoPais", t101.prefijoPais);
        //                    writer.WriteElementString("Numero", t101.numero);
        //                }
        //                writer.WriteEndElement(); // Telefono

        //                writer.WriteElementString("IndicadorTipoDireccion", t101.indicadorTipoDireccion);
        //            }
        //            writer.WriteEndElement(); // Cliente

        //            writer.WriteStartElement("OtrosdatosCliente");
        //            {
        //                writer.WriteElementString("IndicativoDeDireccionExterna", t101.indicativoDeDireccionExterna);
        //                writer.WriteElementString("Linea1DeLaDireccionExterna", t101.calleCliente + " " + t101.numeroFincaCliente);
        //                writer.WriteElementString("Linea2DeLaDireccionExterna", municipio.DesMunicipio(t101.municipioCliente));
        //                writer.WriteElementString("Linea3DeLaDireccionExterna", t101.codPostalCliente + ", "
        //                    + provincias.DesProvincia(t101.codPostalCliente));
        //                writer.WriteElementString("Linea4DeLaDireccionExterna", t101.linea4DeLaDireccionExterna);
        //                writer.WriteElementString("Idioma", t101.idioma);
        //            }
        //            writer.WriteEndElement(); // OtrosdatosCliente
        //        }
        //        writer.WriteEndElement(); // RegistrodePuntodeSuministro

        //        writer.WriteStartElement("ActivacionCambiodeComercializadorSinCambios");
        //        {
        //            writer.WriteStartElement("DatosActivacion");
        //            {
        //                writer.WriteElementString("Fecha", t105.fechaActivacion.ToString("yyyy-MM-dd"));
        //            }
        //            writer.WriteEndElement(); // DatosActivacion

        //            writer.WriteStartElement("Contrato");
        //            {
        //                writer.WriteStartElement("IdContrato");
        //                {
        //                    writer.WriteElementString("CodContrato", t101.codContrato);
        //                }
        //                writer.WriteEndElement(); // IdContrato

        //                writer.WriteElementString("TipoAutoconsumo", t105.tipoAutoconsumo);
        //                writer.WriteElementString("TipoContratoATR", t105.tipoContratoATR);

        //                writer.WriteStartElement("CondicionesContractuales");
        //                {
        //                    writer.WriteElementString("TarifaATR", t101.tarifaATR);
        //                    writer.WriteElementString("PeriodicidadFacturacion", t105.periodicidadFacturacion);
        //                    writer.WriteElementString("TipodeTelegestion", t105.tipodeTelegestion);

        //                    writer.WriteStartElement("PotenciasContratadas");
        //                    {
        //                        for (int i = 1; i < 7; i++)
        //                        {
        //                            if (t101.potenciaPeriodo[i] > 0)
        //                            {
        //                                writer.WriteStartElement("Potencia");
        //                                writer.WriteAttributeString("Periodo", i.ToString());
        //                                writer.WriteString(t101.potenciaPeriodo[i].ToString());
        //                                writer.WriteEndElement();
        //                                //writer.WriteElementString("Potencia", "Periodo=" + i, xml.potenciaPeriodo[i].ToString());
        //                            }

        //                            //writer.WriteElementString("Potencia Periodo=" + @"""" + (i + 1) + @"""", xml.potenciaPeriodo[i].ToString());
        //                        }

        //                    }
        //                    writer.WriteEndElement(); // PotenciasContratadas

        //                    writer.WriteElementString("MarcaMedidaConPerdidas", t105.marcaMedidaConPerdidas);
        //                    writer.WriteElementString("TensionDelSuministro", t105.tensionDelSuministro.ToString());
        //                }
        //                writer.WriteEndElement(); // CondicionesContractuales
        //            }
        //            writer.WriteEndElement(); // Contrato

        //            writer.WriteStartElement("PuntosDeMedida");
        //            {
        //                for (int i = 0; i < t105.lista_PM.Count; i++)
        //                {
        //                    writer.WriteStartElement("PuntoDeMedida");
        //                    writer.WriteElementString("CodPM", t105.lista_PM[i].codPM);
        //                    writer.WriteElementString("TipoMovimiento", t105.lista_PM[i].tipoMovimiento);
        //                    writer.WriteElementString("TipoPM", t105.lista_PM[i].tipoPM);
        //                    writer.WriteElementString("ModoLectura", t105.lista_PM[i].modoLectura);
        //                    writer.WriteElementString("Funcion", t105.lista_PM[i].funcion);

        //                    writer.WriteElementString("TelefonoTelemedida", t105.lista_PM[i].telefonoTelemedida);
        //                    writer.WriteElementString("TensionPM", t105.lista_PM[i].tensionPM);
        //                    writer.WriteElementString("FechaVigor", t105.lista_PM[i].fechaVigor.ToString("yyyy-MM-dd"));
        //                    writer.WriteElementString("FechaAlta", t105.lista_PM[i].fechaAlta.ToString("yyyy-MM-dd"));
        //                    writer.WriteElementString("FechaBaja", t105.lista_PM[i].fechaBaja.ToString("yyyy-MM-dd"));

        //                    writer.WriteStartElement("Aparatos");
        //                    {
        //                        for (int j = 0; j < t105.lista_PM[i].lista_aparatos.Count; j++)
        //                        {
        //                            writer.WriteStartElement("Aparato");
        //                            {
        //                                writer.WriteStartElement("ModeloAparato");
        //                                {
        //                                    writer.WriteElementString("TipoAparato", t105.lista_PM[i].lista_aparatos[j].tipoAparato);
        //                                    writer.WriteElementString("MarcaAparato", t105.lista_PM[i].lista_aparatos[j].marcaAparato);
        //                                    writer.WriteElementString("ModeloMarca", t105.lista_PM[i].lista_aparatos[j].modeloMarca);
        //                                }
        //                                writer.WriteEndElement(); // ModeloAparato

        //                                writer.WriteElementString("TipoMovimiento", t105.lista_PM[i].lista_aparatos[j].tipoMovimiento);
        //                                writer.WriteElementString("TipoEquipoMedida", t105.lista_PM[i].lista_aparatos[j].tipoEquipoMedida);
        //                                writer.WriteElementString("TipoPropiedadAparato", t105.lista_PM[i].lista_aparatos[j].tipoPropiedadAparato);
        //                                writer.WriteElementString("TipoDHEdM", t105.lista_PM[i].lista_aparatos[j].tipoDHEdM);
        //                                writer.WriteElementString("ModoMedidaPotencia", t105.lista_PM[i].lista_aparatos[j].modoMedidaPotencia);
        //                                writer.WriteElementString("CodPrecinto", t105.lista_PM[i].lista_aparatos[j].codPrecinto);

        //                                writer.WriteStartElement("DatosAparato");
        //                                {
        //                                    writer.WriteElementString("PeriodoFabricacion", t105.lista_PM[i].lista_aparatos[j].periodoFabricacion);
        //                                    writer.WriteElementString("NumeroSerie", t105.lista_PM[i].lista_aparatos[j].numeroSerie);
        //                                    writer.WriteElementString("FuncionAparato", t105.lista_PM[i].lista_aparatos[j].funcionAparato);
        //                                    writer.WriteElementString("NumIntegradores", t105.lista_PM[i].lista_aparatos[j].numIntegradores);
        //                                    writer.WriteElementString("ConstanteEnergia", t105.lista_PM[i].lista_aparatos[j].constanteEnergia);
        //                                    writer.WriteElementString("ConstanteMaximetro", t105.lista_PM[i].lista_aparatos[j].constanteMaximetro);
        //                                    writer.WriteElementString("RuedasEnteras", t105.lista_PM[i].lista_aparatos[j].ruedasEnteras);
        //                                    writer.WriteElementString("RuedasDecimales", t105.lista_PM[i].lista_aparatos[j].ruedasDecimales);

        //                                }
        //                                writer.WriteEndElement(); // DatosAparato

        //                                if (t105.lista_PM[i].lista_aparatos[j].listaMedidas.Count > 0)
        //                                {
        //                                    writer.WriteStartElement("Medidas");
        //                                    {
        //                                        foreach (EndesaEntity.contratacion.xxi.XML_Medidas p in t105.lista_PM[i].lista_aparatos[j].listaMedidas)
        //                                        {
        //                                            writer.WriteStartElement("Medida");
        //                                            writer.WriteElementString("TipoDHEdM", p.tipoDHEdM);
        //                                            writer.WriteElementString("Periodo", p.periodo);
        //                                            writer.WriteElementString("MagnitudMedida", p.magnitudMedida);
        //                                            writer.WriteElementString("Procedencia", p.procedencia);
        //                                            writer.WriteElementString("UltimaLecturaFirme", p.ultimaLecturaFirme);
        //                                            writer.WriteElementString("FechaLecturaFirme", p.fechaLecturaFirme.ToString("yyyy-MM-dd"));
        //                                            writer.WriteEndElement(); // Medida
        //                                        }
        //                                    }
        //                                    writer.WriteEndElement(); // Medidas

        //                                }
        //                            }
        //                            writer.WriteEndElement(); // Aparato

        //                        }
        //                    }
        //                    writer.WriteEndElement(); // Aparatos

        //                }
        //                writer.WriteEndElement(); // PuntoDeMedida
        //            }
        //            writer.WriteEndElement(); // PuntosDeMedida

        //        }
        //        writer.WriteEndElement(); // ActivacionCambiodeComercializadorSinCambios                  

        //    }


        //    writer.WriteEndElement(); //VueltasaCUR
        //    writer.WriteEndDocument();

        //    writer.Close();


        //}

        //public void CreaXML_t105_A302_A305(FileInfo file, EndesaEntity.contratacion.xxi.XML_Datos t101, EndesaEntity.contratacion.xxi.XML_Datos t105)
        //{

        //    string rr = "*****************";


        //    #region RegistrodePuntodeSuministro
        //    if (t101.pais == "")
        //        t101.pais = t101.cups.Substring(0, 2);
        //    if (t101.provincia == "")
        //        t101.provincia = "NN";
        //    if (t101.municipio == "")
        //        t101.municipio = "NNNNN";
        //    if (t101.poblacion == "")
        //        t101.poblacion = "NNNNNNNNN";
        //    if (t101.descripcionPoblacion == "")
        //        t101.descripcionPoblacion = rr;
        //    if (t101.tipoVia == "")
        //        t101.tipoVia = "**";
        //    if (t101.codPostal == "")
        //        t101.codPostal = "NNNNN";
        //    if (t101.calle == "")
        //        t101.calle = rr;
        //    if (t101.numeroFinca == "")
        //        t101.numeroFinca = rr;
        //    if (t101.piso == "")
        //        t101.piso = rr;

        //    if (t105.indicativoDeInterrumpibilidad == "")
        //        t105.indicativoDeInterrumpibilidad = "N";

        //    if (t101.tipoIdentificador == "")
        //        t101.tipoIdentificador = "NI";
        //    if (t101.identificador == "")
        //        t101.identificador = rr;
        //    if (t101.tipoPersona == "")
        //        t101.tipoPersona = "J";
        //    //if (t101.cliente_NombreDePila == "")
        //    //    t101.cliente_NombreDePila = rr;
        //    //if (t101.cliente_PrimerApellido == "")
        //    //    t101.cliente_PrimerApellido = rr;

        //    if (t101.descripcionPoblacion == "")
        //        t101.descripcionPoblacion = rr;

        //    #region "DireccionPS
        //    {
        //        if (t101.direccionPS_Pais == "")
        //            t101.direccionPS_Pais = t101.cups.Substring(0, 2);

        //        if (t101.direccionPS_Provincia == "") 
        //            t101.direccionPS_Provincia = "NN";

        //        if (t101.direccionPS_Municipio == "") 
        //            t101.direccionPS_Municipio = "NNNNN";

        //        if(t101.direccionPS_Poblacion == "")
        //            t101.direccionPS_Poblacion = "NNNNNNNNN";

        //        if (t101.direccionPS_CodPostal == "")
        //            t101.direccionPS_CodPostal = "NNNNN";

        //        if (t101.direccionPS_Calle == "")
        //            t101.direccionPS_Calle = rr;

        //        if (t101.direccionPS_NumeroFinca == "")
        //            t101.direccionPS_NumeroFinca = "SN";                                               

        //        if (t101.direccionPS_AclaradorFinca == "" || t101.direccionPS_AclaradorFinca == null)
        //            t101.direccionPS_AclaradorFinca = rr;


        //    }
        //    #endregion

        //    #region Direccion Cliente
        //    {
        //        if (t101.paisCliente == "")
        //            t101.paisCliente = t101.cups.Substring(0, 2);

        //        if (t101.provinciaCliente == "")
        //            t101.provinciaCliente = "NN";

        //        if (t101.municipioCliente == "")
        //            t101.municipioCliente = "NNNNN";

        //        if (t101.poblacionCliente == "")
        //            t101.poblacionCliente = "NNNNNNNNN"; 
                               
                
        //        if (t101.codPostalCliente == "")
        //            t101.codPostalCliente = "NNNNN";

        //        if (t101.calleCliente == "")
        //            t101.calleCliente = rr;

        //        if (t101.numeroFincaCliente == "")
        //            t101.numeroFincaCliente = "SN";

        //        if (t101.aclaradorFinca == "")
        //            t101.aclaradorFinca = rr;

        //        if (t101.prefijoPais == "" || t101.prefijoPais == null)
        //            t101.prefijoPais = "34";

        //        if (t101.numero == "" || t101.numero == null)
        //            t101.numero = "000000000";

              


        //    }
        //    #endregion


        //    if (t101.razonSocial == "")
        //        t101.razonSocial = rr;



        //    if (t101.indicativoDeDireccionExterna == "S")
        //    {
        //        if (t101.linea1DeLaDireccionExterna == "")
        //            t101.linea1DeLaDireccionExterna = rr;
        //        if (t101.linea2DeLaDireccionExterna == "")
        //            t101.linea2DeLaDireccionExterna = rr;
        //        if (t101.linea3DeLaDireccionExterna == "")
        //            t101.linea3DeLaDireccionExterna = rr;
        //        if (t101.linea4DeLaDireccionExterna == "")
        //            t101.linea4DeLaDireccionExterna = rr;
        //    }



        //    #endregion


        //    string url = "http://xmlns.endesa.com/wsdl/FormatoCUR/";

        //    XmlTextWriter writer = new XmlTextWriter(file.FullName, Encoding.UTF8);


        //    writer.WriteStartDocument();
        //    writer.WriteStartElement("VueltasaCUR", url);
        //    {
        //        writer.WriteStartElement("Cabecera");
        //        {
        //            writer.WriteElementString("CodigoREEEmpresaEmisora", t101.codigoREEEmpresaEmisora);
        //            writer.WriteElementString("CodigoREEEmpresaDestino", t101.codigoREEEmpresaDestino);
        //            writer.WriteElementString("CodigoDelProceso", t105.codigoDelProceso);
        //            writer.WriteElementString("CodigoDePaso", t105.codigoDePaso);
        //            writer.WriteElementString("CodigoDeSolicitud", t105.codigoDeSolicitud);
        //            writer.WriteElementString("SecuencialDeSolicitud", t105.secuencialDeSolicitud);
        //            writer.WriteElementString("FechaSolicitud", t101.fechaSolicitud.ToString("yyyy-MM-ddTHH:mm:ss"));
        //            writer.WriteElementString("CUPS", t101.cups);
        //        }
        //        writer.WriteEndElement(); // Cabecera

        //        writer.WriteStartElement("RegistrodePuntodeSuministro");
        //        {
        //            writer.WriteStartElement("Suministro");
        //            {
        //                writer.WriteElementString("CNAE", t101.cnae.ToString());
        //                writer.WriteElementString("PotenciaExtension", t105.potenciaExtension.ToString());
        //                writer.WriteElementString("PotenciaDeAcceso", t105.potenciaDeAcceso.ToString());                        
        //                writer.WriteElementString("IndicativoDeInterrumpibilidad", t105.indicativoDeInterrumpibilidad);

        //                writer.WriteStartElement("DireccionPS");
        //                {
        //                    writer.WriteElementString("Pais", t101.direccionPS_Pais);                          
        //                    writer.WriteElementString("Provincia", t101.direccionPS_Provincia);                            
        //                    writer.WriteElementString("Municipio",t101.direccionPS_Municipio);                            
        //                    writer.WriteElementString("Poblacion", t101.direccionPS_Poblacion);                            
        //                    writer.WriteElementString("CodPostal", t101.direccionPS_CodPostal);
        //                    writer.WriteElementString("Calle", t101.direccionPS_Calle);
        //                    writer.WriteElementString("NumeroFinca", t101.direccionPS_NumeroFinca);                                                        
        //                    writer.WriteElementString("AclaradorFinca", t101.direccionPS_AclaradorFinca);
        //                }
        //                writer.WriteEndElement(); // DireccionPS

        //            }
        //            writer.WriteEndElement(); // Suministro

        //            writer.WriteStartElement("Cliente");
        //            {
        //                writer.WriteStartElement("IdCliente");
        //                {
        //                    writer.WriteElementString("TipoIdentificador", t101.tipoIdentificador);
        //                    writer.WriteElementString("Identificador", t101.identificador);
        //                    writer.WriteElementString("TipoPersona", t101.tipoPersona);
        //                }
        //                writer.WriteEndElement(); // IdCliente

        //                writer.WriteStartElement("Nombre");
        //                {
        //                    writer.WriteElementString("RazonSocial", t101.razonSocial);
        //                }
        //                writer.WriteEndElement(); // Nombre

        //                writer.WriteStartElement("Telefono");
        //                {
        //                    writer.WriteElementString("PrefijoPais", t101.prefijoPais);
        //                    writer.WriteElementString("Numero", t101.numero);
        //                }
        //                writer.WriteEndElement(); // Telefono

        //                writer.WriteElementString("IndicadorTipoDireccion", t101.indicadorTipoDireccion);

        //                writer.WriteStartElement("Direccion");
        //                {
        //                    writer.WriteElementString("Pais", t101.paisCliente);                            
        //                    writer.WriteElementString("Provincia", t101.provinciaCliente);                            
        //                    writer.WriteElementString("Municipio", t101.municipioCliente);                            
        //                    writer.WriteElementString("Poblacion", t101.poblacionCliente);                            
        //                    writer.WriteElementString("CodPostal", t101.codPostalCliente);
        //                    writer.WriteElementString("Calle", t101.calleCliente);
        //                    writer.WriteElementString("NumeroFinca", t101.numeroFincaCliente);
        //                    writer.WriteElementString("AclaradorFinca", t101.aclaradorFinca);
        //                }
        //                writer.WriteEndElement(); // Direccion
        //            }
        //            writer.WriteEndElement(); // Cliente

        //            writer.WriteStartElement("OtrosdatosCliente");
        //            {
        //                //writer.WriteElementString("IndicativoDeDireccionExterna", t101.indicativoDeDireccionExterna);
        //                //writer.WriteElementString("Linea1DeLaDireccionExterna", t101.calleCliente + " " + t101.numeroFincaCliente);
        //                //writer.WriteElementString("Linea2DeLaDireccionExterna", municipio.DesMunicipio(t101.municipioCliente));
        //                //writer.WriteElementString("Linea3DeLaDireccionExterna",t101.codPostalCliente + ", " 
        //                //    + provincias.DesProvincia(t101.codPostalCliente));
        //                //writer.WriteElementString("Linea4DeLaDireccionExterna", t101.linea4DeLaDireccionExterna);

        //                //writer.WriteElementString("IndicativoDeDireccionExterna", t101.indicativoDeDireccionExterna);
        //                //writer.WriteElementString("Linea1DeLaDireccionExterna", t101.linea1DeLaDireccionExterna);
        //                //writer.WriteElementString("Linea2DeLaDireccionExterna", t101.linea2DeLaDireccionExterna);
        //                //writer.WriteElementString("Linea3DeLaDireccionExterna", t101.linea3DeLaDireccionExterna);
        //                //writer.WriteElementString("Linea4DeLaDireccionExterna", t101.linea4DeLaDireccionExterna);
        //                writer.WriteElementString("Idioma", t101.idioma);
        //            }
        //            writer.WriteEndElement(); // OtrosdatosCliente
        //        }
        //        writer.WriteEndElement(); // RegistrodePuntodeSuministro

        //        writer.WriteStartElement("ActivacionCambiodeComercializadorSinCambios");
        //        {
        //            writer.WriteStartElement("DatosActivacion");
        //            {
        //                writer.WriteElementString("Fecha", t105.fechaActivacion.ToString("yyyy-MM-dd"));
        //            }
        //            writer.WriteEndElement(); // DatosActivacion

        //            writer.WriteStartElement("Contrato");
        //            {
        //                writer.WriteStartElement("IdContrato");
        //                {
        //                    writer.WriteElementString("CodContrato", t105.codContrato);
        //                }
        //                writer.WriteEndElement(); // IdContrato

        //                writer.WriteElementString("TipoAutoconsumo", t105.tipoAutoconsumo);
        //                writer.WriteElementString("TipoContratoATR", t105.tipoContratoATR);

        //                writer.WriteStartElement("CondicionesContractuales");
        //                {
        //                    writer.WriteElementString("TarifaATR", t101.tarifaATR);
        //                    writer.WriteElementString("PeriodicidadFacturacion", t105.periodicidadFacturacion);
        //                    writer.WriteElementString("TipodeTelegestion", t105.tipodeTelegestion);

        //                    writer.WriteStartElement("PotenciasContratadas");
        //                    {
        //                        for (int i = 1; i < 7; i++)
        //                        {
        //                            if (t101.potenciaPeriodo[i] > 0)
        //                            {
        //                                writer.WriteStartElement("Potencia");
        //                                writer.WriteAttributeString("Periodo", i.ToString());
        //                                writer.WriteString(t101.potenciaPeriodo[i].ToString());
        //                                writer.WriteEndElement();
        //                                //writer.WriteElementString("Potencia", "Periodo=" + i, xml.potenciaPeriodo[i].ToString());
        //                            }

        //                            //writer.WriteElementString("Potencia Periodo=" + @"""" + (i + 1) + @"""", xml.potenciaPeriodo[i].ToString());
        //                        }

        //                    }
        //                    writer.WriteEndElement(); // PotenciasContratadas

        //                    writer.WriteElementString("MarcaMedidaConPerdidas", t105.marcaMedidaConPerdidas);
        //                    writer.WriteElementString("TensionDelSuministro", t105.tensionDelSuministro.ToString());
        //                    writer.WriteElementString("VAsTrafo", t105.vasTrafo.ToString());
        //                }
        //                writer.WriteEndElement(); // CondicionesContractuales
        //            }
        //            writer.WriteEndElement(); // Contrato

        //            writer.WriteStartElement("PuntosDeMedida");
        //            {
        //                for (int i = 0; i < t105.lista_PM.Count; i++)
        //                {
        //                    writer.WriteStartElement("PuntoDeMedida");
        //                    writer.WriteElementString("CodPM", t105.lista_PM[i].codPM);
        //                    writer.WriteElementString("TipoMovimiento", t105.lista_PM[i].tipoMovimiento);
        //                    writer.WriteElementString("TipoPM", t105.lista_PM[i].tipoPM);
        //                    writer.WriteElementString("ModoLectura", t105.lista_PM[i].modoLectura);
        //                    writer.WriteElementString("Funcion", t105.lista_PM[i].funcion);

        //                    writer.WriteElementString("TelefonoTelemedida", t105.lista_PM[i].telefonoTelemedida);
        //                    writer.WriteElementString("TensionPM", t105.lista_PM[i].tensionPM);
        //                    writer.WriteElementString("FechaVigor", t105.lista_PM[i].fechaVigor.ToString("yyyy-MM-dd"));
        //                    writer.WriteElementString("FechaAlta", t105.lista_PM[i].fechaAlta.ToString("yyyy-MM-dd"));
        //                    //writer.WriteElementString("FechaBaja", t105.lista_PM[i].fechaBaja.ToString("yyyy-MM-dd"));

        //                    writer.WriteStartElement("Aparatos");
        //                    {
        //                        for (int j = 0; j < t105.lista_PM[i].lista_aparatos.Count; j++)
        //                        {
        //                            writer.WriteStartElement("Aparato");
        //                            {
        //                                writer.WriteStartElement("ModeloAparato");
        //                                {
        //                                    writer.WriteElementString("TipoAparato", t105.lista_PM[i].lista_aparatos[j].tipoAparato);
        //                                    writer.WriteElementString("MarcaAparato", t105.lista_PM[i].lista_aparatos[j].marcaAparato);
        //                                    writer.WriteElementString("ModeloMarca", t105.lista_PM[i].lista_aparatos[j].modeloMarca);
        //                                }
        //                                writer.WriteEndElement(); // ModeloAparato

        //                                writer.WriteElementString("TipoMovimiento", t105.lista_PM[i].lista_aparatos[j].tipoMovimiento);
        //                                writer.WriteElementString("TipoEquipoMedida", t105.lista_PM[i].lista_aparatos[j].tipoEquipoMedida);
        //                                writer.WriteElementString("TipoPropiedadAparato", t105.lista_PM[i].lista_aparatos[j].tipoPropiedadAparato);
        //                                writer.WriteElementString("TipoDHEdM", t105.lista_PM[i].lista_aparatos[j].tipoDHEdM);
        //                                writer.WriteElementString("ModoMedidaPotencia", t105.lista_PM[i].lista_aparatos[j].modoMedidaPotencia);
        //                                writer.WriteElementString("CodPrecinto", t105.lista_PM[i].lista_aparatos[j].codPrecinto);

        //                                writer.WriteStartElement("DatosAparato");
        //                                {
        //                                    writer.WriteElementString("PeriodoFabricacion", t105.lista_PM[i].lista_aparatos[j].periodoFabricacion);
        //                                    writer.WriteElementString("NumeroSerie", t105.lista_PM[i].lista_aparatos[j].numeroSerie);
        //                                    writer.WriteElementString("FuncionAparato", t105.lista_PM[i].lista_aparatos[j].funcionAparato);
        //                                    writer.WriteElementString("NumIntegradores", t105.lista_PM[i].lista_aparatos[j].numIntegradores);
        //                                    writer.WriteElementString("ConstanteEnergia", t105.lista_PM[i].lista_aparatos[j].constanteEnergia);
        //                                    writer.WriteElementString("ConstanteMaximetro", t105.lista_PM[i].lista_aparatos[j].constanteMaximetro);
        //                                    writer.WriteElementString("RuedasEnteras", t105.lista_PM[i].lista_aparatos[j].ruedasEnteras);
        //                                    writer.WriteElementString("RuedasDecimales", t105.lista_PM[i].lista_aparatos[j].ruedasDecimales);

        //                                }
        //                                writer.WriteEndElement(); // DatosAparato

        //                                if (t105.lista_PM[i].lista_aparatos[j].listaMedidas.Count > 0)
        //                                {
        //                                    writer.WriteStartElement("Medidas");
        //                                    {
        //                                        foreach (EndesaEntity.contratacion.xxi.XML_Medidas p in t105.lista_PM[i].lista_aparatos[j].listaMedidas)
        //                                        {
        //                                            writer.WriteStartElement("Medida");
        //                                            writer.WriteElementString("TipoDHEdM", p.tipoDHEdM);
        //                                            writer.WriteElementString("Periodo", p.periodo);
        //                                            writer.WriteElementString("MagnitudMedida", p.magnitudMedida);
        //                                            writer.WriteElementString("Procedencia", p.procedencia);
        //                                            writer.WriteElementString("UltimaLecturaFirme", p.ultimaLecturaFirme);
        //                                            writer.WriteElementString("FechaLecturaFirme", p.fechaLecturaFirme.ToString("yyyy-MM-dd"));
        //                                            writer.WriteEndElement(); // Medida
        //                                        }
        //                                    }
        //                                    writer.WriteEndElement(); // Medidas

        //                                }
        //                            }
        //                            writer.WriteEndElement(); // Aparato

        //                        }
        //                    }
        //                    writer.WriteEndElement(); // Aparatos

        //                }
        //                writer.WriteEndElement(); // PuntoDeMedida
        //            }
        //            writer.WriteEndElement(); // PuntosDeMedida

        //        }
        //        writer.WriteEndElement(); // ActivacionCambiodeComercializadorSinCambios                  

        //    }


        //    writer.WriteEndElement(); //VueltasaCUR
        //    writer.WriteEndDocument();

        //    writer.Close();


        //}

        public bool CreaXML_t105_A302_A305_v2(FileInfo file, EndesaEntity.contratacion.xxi.XML_Datos t101, 
            EndesaEntity.contratacion.xxi.XML_Datos t105)
        {

            string rr = "*****************";

            try
            {


                #region RegistrodePuntodeSuministro
                if (t101.pais == "")
                    t101.pais = t101.cups.Substring(0, 2);
                if (t101.provincia == "")
                    t101.provincia = "NN";
                if (t101.municipio == "")
                    t101.municipio = "NNNNN";
                if (t101.poblacion == "")
                    t101.poblacion = "NNNNNNNNN";
                if (t101.descripcionPoblacion == "")
                    t101.descripcionPoblacion = rr;
                if (t101.tipoVia == "")
                    t101.tipoVia = "**";
                if (t101.codPostal == "")
                    t101.codPostal = "NNNNN";
                if (t101.calle == "")
                    t101.calle = rr;
                if (t101.numeroFinca == "")
                    t101.numeroFinca = rr;
                if (t101.piso == "")
                    t101.piso = rr;

                if (t105.indicativoDeInterrumpibilidad == "")
                    t105.indicativoDeInterrumpibilidad = "N";

                if (t101.tipoIdentificador == "")
                    t101.tipoIdentificador = "NI";
                if (t101.identificador == "")
                    t101.identificador = rr;
                if (t101.tipoPersona == "")
                    t101.tipoPersona = "J";
                //if (t101.cliente_NombreDePila == "")
                //    t101.cliente_NombreDePila = rr;
                //if (t101.cliente_PrimerApellido == "")
                //    t101.cliente_PrimerApellido = rr;

                if (t101.descripcionPoblacion == "")
                    t101.descripcionPoblacion = rr;

                #region "DireccionPS
                {
                    if (t101.direccionPS_Pais == "")
                        t101.direccionPS_Pais = t101.cups.Substring(0, 2);

                    if (t101.direccionPS_Provincia == "")
                        t101.direccionPS_Provincia = "NN";

                    if (t101.direccionPS_Municipio == "")
                        t101.direccionPS_Municipio = "NNNNN";

                    if (t101.direccionPS_Poblacion == "")
                        t101.direccionPS_Poblacion = "NNNNNNNNN";

                    if (t101.direccionPS_CodPostal == "")
                        t101.direccionPS_CodPostal = "NNNNN";

                    if (t101.direccionPS_Calle == "")
                        t101.direccionPS_Calle = rr;

                    if (t101.direccionPS_NumeroFinca == "")
                        t101.direccionPS_NumeroFinca = "SN";

                    if (t101.direccionPS_AclaradorFinca == "" || t101.direccionPS_AclaradorFinca == null)
                        t101.direccionPS_AclaradorFinca = rr;


                }
                #endregion

                #region Direccion Cliente
                {
                    if (t101.paisCliente == "")
                        t101.paisCliente = t101.cups.Substring(0, 2);

                    if (t101.provinciaCliente == "")
                        t101.provinciaCliente = "NN";

                    if (t101.municipioCliente == "")
                        t101.municipioCliente = "NNNNN";

                    if (t101.poblacionCliente == "")
                        t101.poblacionCliente = "NNNNNNNNN";


                    if (t101.codPostalCliente == "")
                        t101.codPostalCliente = "NNNNN";

                    if (t101.calleCliente == "")
                        t101.calleCliente = rr;

                    if (t101.numeroFincaCliente == "")
                        t101.numeroFincaCliente = "SN";

                    if (t101.aclaradorFinca == "")
                        t101.aclaradorFinca = rr;

                    if (t101.prefijoPais == "" || t101.prefijoPais == null)
                        t101.prefijoPais = "34";

                    if (t101.numero == "" || t101.numero == null)
                        t101.numero = "000000000";




                }
                #endregion


                if (t101.razonSocial == "")
                    t101.razonSocial = rr;



                if (t101.indicativoDeDireccionExterna == "S")
                {
                    if (t101.linea1DeLaDireccionExterna == "")
                        t101.linea1DeLaDireccionExterna = rr;
                    if (t101.linea2DeLaDireccionExterna == "")
                        t101.linea2DeLaDireccionExterna = rr;
                    if (t101.linea3DeLaDireccionExterna == "")
                        t101.linea3DeLaDireccionExterna = rr;
                    if (t101.linea4DeLaDireccionExterna == "")
                        t101.linea4DeLaDireccionExterna = rr;
                }



                #endregion

                EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeT105 _t105 =
                   new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeT105();

                _t105.Cabecera.CodigoREEEmpresaEmisora = t101.codigoREEEmpresaEmisora;
                _t105.Cabecera.CodigoREEEmpresaDestino = t101.codigoREEEmpresaDestino;
                _t105.Cabecera.CodigoDelProceso = "T1";
                _t105.Cabecera.CodigoDePaso = "05";
                _t105.Cabecera.CodigoDeSolicitud = t105.codigoDeSolicitud;
                _t105.Cabecera.SecuencialDeSolicitud = t105.secuencialDeSolicitud;
                _t105.Cabecera.FechaSolicitud = t101.fechaSolicitud.ToString("yyyy-MM-ddTHH:mm:ss");
                _t105.Cabecera.CUPS = t101.cups;

                _t105.ActivacionTraspasoCOR.DatosActivacion.fechaActivacion = t105.fechaActivacion.ToString("yyyy-MM-dd");
                _t105.ActivacionTraspasoCOR.DatosActivacion.enServicio = "S";

                IdContrato idcontrato = new IdContrato();
                idcontrato.CodContrato = t105.codContrato;

                _t105.ActivacionTraspasoCOR.Contrato.IdContrato = idcontrato;
                _t105.ActivacionTraspasoCOR.Contrato.TipoAutoconsumo = t105.tipoAutoconsumo;
                _t105.ActivacionTraspasoCOR.Contrato.TipoContratoATR = t105.tipoContratoATR;
                _t105.ActivacionTraspasoCOR.Contrato.CondicionesContractuales.TarifaATR = t105.tarifaATR;
                _t105.ActivacionTraspasoCOR.Contrato.CondicionesContractuales.PeriodicidadFacturacion = t105.periodicidadFacturacion;
                _t105.ActivacionTraspasoCOR.Contrato.CondicionesContractuales.TipoTelegestion = t105.tipodeTelegestion;

                EndesaEntity.cnmc.V21_2019_12_17.PotenciasContratadas potencias_contratadas =
                   new EndesaEntity.cnmc.V21_2019_12_17.PotenciasContratadas();

                EndesaEntity.cnmc.V21_2019_12_17.Potencia potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "1";
                potencia.potencia = t105.potenciaPeriodo[1].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "2";
                potencia.potencia = t105.potenciaPeriodo[2].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "3";
                potencia.potencia = t105.potenciaPeriodo[3].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "4";
                potencia.potencia = t105.potenciaPeriodo[4].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "5";
                potencia.potencia = t105.potenciaPeriodo[5].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "6";
                potencia.potencia = t105.potenciaPeriodo[6].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                _t105.ActivacionTraspasoCOR.Contrato.CondicionesContractuales.PotenciasContratadas = potencias_contratadas;
               

                _t105.ActivacionTraspasoCOR.Contrato.CondicionesContractuales.MarcaMedidaConPerdidas = t105.marcaMedidaConPerdidas;

                if(t101.tensionDelSuministro > 0)
                    _t105.ActivacionTraspasoCOR.Contrato.CondicionesContractuales.TensionDelSuministro = 
                        t105.tensionDelSuministro.ToString();

                if(t105.vasTrafo > 0)
                    _t105.ActivacionTraspasoCOR.Contrato.CondicionesContractuales.VAsTrafo = 
                        t105.vasTrafo.ToString();


                List<PuntoDeMedida> lista_puntos_medida = new List<PuntoDeMedida>();


                // Puntos de Medida
                for (int i = 0; i < t105.lista_PM.Count; i++)
                {
                    EndesaEntity.cnmc.V21_2019_12_17.PuntoDeMedida ps = new EndesaEntity.cnmc.V21_2019_12_17.PuntoDeMedida();
                    ps.CodPM = t105.lista_PM[i].codPM;
                    ps.TipoMovimiento = t105.lista_PM[i].tipoMovimiento;
                    ps.TipoPM = t105.lista_PM[i].tipoPM;
                    ps.ModoLectura = t105.lista_PM[i].modoLectura;
                    ps.Funcion = t105.lista_PM[i].funcion;
                    ps.TelefonoTelemedida = t105.lista_PM[i].telefonoTelemedida;
                    ps.TensionPM = t105.lista_PM[i].tensionPM;
                    ps.FechaVigor = t105.lista_PM[i].fechaVigor.ToString("yyyy-MM-dd");
                    ps.FechaAlta = t105.lista_PM[i].fechaAlta.ToString("yyyy-MM-dd");

                    for (int j = 0; j < t105.lista_PM[i].lista_aparatos.Count; j++)
                    {
                        EndesaEntity.cnmc.V21_2019_12_17.Aparato aparato =
                            new EndesaEntity.cnmc.V21_2019_12_17.Aparato();

                        aparato.ModeloAparato.TipoAparato = t105.lista_PM[i].lista_aparatos[j].tipoAparato;
                        aparato.ModeloAparato.MarcaAparato = t105.lista_PM[i].lista_aparatos[j].marcaAparato;
                        aparato.ModeloAparato.ModeloMarca = t105.lista_PM[i].lista_aparatos[j].modeloMarca;

                        aparato.TipoMovimiento = t105.lista_PM[i].lista_aparatos[j].tipoMovimiento;
                        aparato.TipoEquipoMedida = t105.lista_PM[i].lista_aparatos[j].tipoEquipoMedida;
                        aparato.TipoPropiedadAparato = t105.lista_PM[i].lista_aparatos[j].tipoPropiedadAparato;
                        aparato.TipoDHEdM = t105.lista_PM[i].lista_aparatos[j].tipoDHEdM;
                        aparato.ModoMedidaPotencia = t105.lista_PM[i].lista_aparatos[j].modoMedidaPotencia;
                        //aparato.LecturaDirecta = t105.lista_PM[i].lista_aparatos[j].lecturaDirecta;
                        aparato.CodPrecinto = t105.lista_PM[i].lista_aparatos[j].codPrecinto;

                        EndesaEntity.cnmc.V21_2019_12_17.DatosAparato datosAparato =
                            new EndesaEntity.cnmc.V21_2019_12_17.DatosAparato();

                        datosAparato.PeriodoFabricacion = t105.lista_PM[i].lista_aparatos[j].periodoFabricacion;
                        datosAparato.NumeroSerie = t105.lista_PM[i].lista_aparatos[j].numeroSerie;

                        datosAparato.FuncionAparato = t105.lista_PM[i].lista_aparatos[j].funcionAparato;
                        datosAparato.NumIntegradores = t105.lista_PM[i].lista_aparatos[j].numIntegradores;
                        datosAparato.ConstanteEnergia = t105.lista_PM[i].lista_aparatos[j].constanteEnergia;
                        datosAparato.ConstanteMaximetro = t105.lista_PM[i].lista_aparatos[j].constanteMaximetro;
                        datosAparato.RuedasEnteras = t105.lista_PM[i].lista_aparatos[j].ruedasEnteras;
                        datosAparato.RuedasDecimales = t105.lista_PM[i].lista_aparatos[j].ruedasDecimales;

                        aparato.DatosAparato = datosAparato;

                        ps.Aparatos.Add(aparato);

                        if (t105.lista_PM[i].lista_aparatos[j].listaMedidas.Count > 0)
                        {
                            foreach (EndesaEntity.contratacion.xxi.XML_Medidas p in t105.lista_PM[i].lista_aparatos[j].listaMedidas)
                            {

                                EndesaEntity.cnmc.V21_2019_12_17.Medida medida =
                                    new EndesaEntity.cnmc.V21_2019_12_17.Medida();

                                medida.TipoDHEdM = p.tipoDHEdM;
                                medida.Periodo = p.periodo;
                                medida.MagnitudMedida = p.magnitudMedida;
                                medida.Procedencia = p.procedencia;
                                medida.UltimaLecturaFirme = p.ultimaLecturaFirme;
                                medida.FechaLecturaFirme = p.fechaLecturaFirme.ToString("yyyy-MM-dd");

                                ps.Medidas.Add(medida);
                            }
                        }

                    }

                    lista_puntos_medida.Add(ps);

                    
                }

                _t105.ActivacionTraspasoCOR.PuntosDeMedida.PuntoDeMedida = lista_puntos_medida;

                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                settings.OmitXmlDeclaration = true;
                XmlWriter writer = XmlWriter.Create(file.FullName, settings);
                XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeT105));
                serializer.Serialize(writer, _t105);
                writer.Close();

                #region Comentario
                //string url = "http://xmlns.endesa.com/wsdl/FormatoCUR/";

                //XmlTextWriter writer = new XmlTextWriter(file.FullName, Encoding.UTF8);


                //writer.WriteStartDocument();
                //writer.WriteStartElement("VueltasaCUR", url);
                //{
                //    writer.WriteStartElement("Cabecera");
                //    {
                //        //writer.WriteElementString("CodigoREEEmpresaEmisora", t101.codigoREEEmpresaEmisora);
                //        //writer.WriteElementString("CodigoREEEmpresaDestino", t101.codigoREEEmpresaDestino);
                //        //writer.WriteElementString("CodigoDelProceso", t105.codigoDelProceso);
                //        //writer.WriteElementString("CodigoDePaso", t105.codigoDePaso);
                //        //writer.WriteElementString("CodigoDeSolicitud", t105.codigoDeSolicitud);
                //        //writer.WriteElementString("SecuencialDeSolicitud", t105.secuencialDeSolicitud);
                //        //writer.WriteElementString("FechaSolicitud", t101.fechaSolicitud.ToString("yyyy-MM-ddTHH:mm:ss"));
                //        //writer.WriteElementString("CUPS", t101.cups);
                //    }
                //    writer.WriteEndElement(); // Cabecera

                //    writer.WriteStartElement("RegistrodePuntodeSuministro");
                //    {
                //        writer.WriteStartElement("Suministro");
                //        {
                //            writer.WriteElementString("CNAE", t101.cnae.ToString());
                //            writer.WriteElementString("PotenciaExtension", t105.potenciaExtension.ToString());
                //            writer.WriteElementString("PotenciaDeAcceso", t105.potenciaDeAcceso.ToString());
                //            writer.WriteElementString("IndicativoDeInterrumpibilidad", t105.indicativoDeInterrumpibilidad);

                //            writer.WriteStartElement("DireccionPS");
                //            {
                //                writer.WriteElementString("Pais", t101.direccionPS_Pais);
                //                writer.WriteElementString("Provincia", t101.direccionPS_Provincia);
                //                writer.WriteElementString("Municipio", t101.direccionPS_Municipio);
                //                writer.WriteElementString("Poblacion", t101.direccionPS_Poblacion);
                //                writer.WriteElementString("CodPostal", t101.direccionPS_CodPostal);
                //                writer.WriteElementString("Calle", t101.direccionPS_Calle);
                //                writer.WriteElementString("NumeroFinca", t101.direccionPS_NumeroFinca);
                //                writer.WriteElementString("AclaradorFinca", t101.direccionPS_AclaradorFinca);
                //            }
                //            writer.WriteEndElement(); // DireccionPS

                //        }
                //        writer.WriteEndElement(); // Suministro

                //        writer.WriteStartElement("Cliente");
                //        {
                //            writer.WriteStartElement("IdCliente");
                //            {
                //                writer.WriteElementString("TipoIdentificador", t101.tipoIdentificador);
                //                writer.WriteElementString("Identificador", t101.identificador);
                //                writer.WriteElementString("TipoPersona", t101.tipoPersona);
                //            }
                //            writer.WriteEndElement(); // IdCliente

                //            writer.WriteStartElement("Nombre");
                //            {
                //                writer.WriteElementString("RazonSocial", t101.razonSocial);
                //            }
                //            writer.WriteEndElement(); // Nombre

                //            writer.WriteStartElement("Telefono");
                //            {
                //                writer.WriteElementString("PrefijoPais", t101.prefijoPais);
                //                writer.WriteElementString("Numero", t101.numero);
                //            }
                //            writer.WriteEndElement(); // Telefono

                //            writer.WriteElementString("IndicadorTipoDireccion", t101.indicadorTipoDireccion);

                //            writer.WriteStartElement("Direccion");
                //            {
                //                writer.WriteElementString("Pais", t101.paisCliente);
                //                writer.WriteElementString("Provincia", t101.provinciaCliente);
                //                writer.WriteElementString("Municipio", t101.municipioCliente);
                //                writer.WriteElementString("Poblacion", t101.poblacionCliente);
                //                writer.WriteElementString("CodPostal", t101.codPostalCliente);
                //                writer.WriteElementString("Calle", t101.calleCliente);
                //                writer.WriteElementString("NumeroFinca", t101.numeroFincaCliente);
                //                writer.WriteElementString("AclaradorFinca", t101.aclaradorFinca);
                //            }
                //            writer.WriteEndElement(); // Direccion
                //        }
                //        writer.WriteEndElement(); // Cliente

                //        writer.WriteStartElement("OtrosdatosCliente");
                //        {
                //            //writer.WriteElementString("IndicativoDeDireccionExterna", t101.indicativoDeDireccionExterna);
                //            //writer.WriteElementString("Linea1DeLaDireccionExterna", t101.calleCliente + " " + t101.numeroFincaCliente);
                //            //writer.WriteElementString("Linea2DeLaDireccionExterna", municipio.DesMunicipio(t101.municipioCliente));
                //            //writer.WriteElementString("Linea3DeLaDireccionExterna",t101.codPostalCliente + ", " 
                //            //    + provincias.DesProvincia(t101.codPostalCliente));
                //            //writer.WriteElementString("Linea4DeLaDireccionExterna", t101.linea4DeLaDireccionExterna);

                //            //writer.WriteElementString("IndicativoDeDireccionExterna", t101.indicativoDeDireccionExterna);
                //            //writer.WriteElementString("Linea1DeLaDireccionExterna", t101.linea1DeLaDireccionExterna);
                //            //writer.WriteElementString("Linea2DeLaDireccionExterna", t101.linea2DeLaDireccionExterna);
                //            //writer.WriteElementString("Linea3DeLaDireccionExterna", t101.linea3DeLaDireccionExterna);
                //            //writer.WriteElementString("Linea4DeLaDireccionExterna", t101.linea4DeLaDireccionExterna);
                //            writer.WriteElementString("Idioma", t101.idioma);
                //        }
                //        writer.WriteEndElement(); // OtrosdatosCliente
                //    }
                //    writer.WriteEndElement(); // RegistrodePuntodeSuministro

                //    writer.WriteStartElement("ActivacionCambiodeComercializadorSinCambios");
                //    {
                //        writer.WriteStartElement("DatosActivacion");
                //        {
                //            writer.WriteElementString("Fecha", t105.fechaActivacion.ToString("yyyy-MM-dd"));
                //        }
                //        writer.WriteEndElement(); // DatosActivacion

                //        writer.WriteStartElement("Contrato");
                //        {
                //            writer.WriteStartElement("IdContrato");
                //            {
                //                writer.WriteElementString("CodContrato", t105.codContrato);
                //            }
                //            writer.WriteEndElement(); // IdContrato

                //            writer.WriteElementString("TipoAutoconsumo", t105.tipoAutoconsumo);
                //            writer.WriteElementString("TipoContratoATR", t105.tipoContratoATR);

                //            writer.WriteStartElement("CondicionesContractuales");
                //            {
                //                writer.WriteElementString("TarifaATR", t101.tarifaATR);
                //                writer.WriteElementString("PeriodicidadFacturacion", t105.periodicidadFacturacion);
                //                writer.WriteElementString("TipodeTelegestion", t105.tipodeTelegestion);

                //                writer.WriteStartElement("PotenciasContratadas");
                //                {
                //                    for (int i = 1; i < 7; i++)
                //                    {
                //                        if (t101.potenciaPeriodo[i] > 0)
                //                        {
                //                            writer.WriteStartElement("Potencia");
                //                            writer.WriteAttributeString("Periodo", i.ToString());
                //                            writer.WriteString(t101.potenciaPeriodo[i].ToString());
                //                            writer.WriteEndElement();
                //                            //writer.WriteElementString("Potencia", "Periodo=" + i, xml.potenciaPeriodo[i].ToString());
                //                        }

                //                        //writer.WriteElementString("Potencia Periodo=" + @"""" + (i + 1) + @"""", xml.potenciaPeriodo[i].ToString());
                //                    }

                //                }
                //                writer.WriteEndElement(); // PotenciasContratadas

                //                writer.WriteElementString("MarcaMedidaConPerdidas", t105.marcaMedidaConPerdidas);
                //                writer.WriteElementString("TensionDelSuministro", t105.tensionDelSuministro.ToString());
                //                writer.WriteElementString("VAsTrafo", t105.vasTrafo.ToString());
                //            }
                //            writer.WriteEndElement(); // CondicionesContractuales
                //        }
                //        writer.WriteEndElement(); // Contrato

                //        writer.WriteStartElement("PuntosDeMedida");
                //        {
                //            for (int i = 0; i < t105.lista_PM.Count; i++)
                //            {
                //                writer.WriteStartElement("PuntoDeMedida");
                //                writer.WriteElementString("CodPM", t105.lista_PM[i].codPM);
                //                writer.WriteElementString("TipoMovimiento", t105.lista_PM[i].tipoMovimiento);
                //                writer.WriteElementString("TipoPM", t105.lista_PM[i].tipoPM);
                //                writer.WriteElementString("ModoLectura", t105.lista_PM[i].modoLectura);
                //                writer.WriteElementString("Funcion", t105.lista_PM[i].funcion);

                //                writer.WriteElementString("TelefonoTelemedida", t105.lista_PM[i].telefonoTelemedida);
                //                writer.WriteElementString("TensionPM", t105.lista_PM[i].tensionPM);
                //                writer.WriteElementString("FechaVigor", t105.lista_PM[i].fechaVigor.ToString("yyyy-MM-dd"));
                //                writer.WriteElementString("FechaAlta", t105.lista_PM[i].fechaAlta.ToString("yyyy-MM-dd"));
                //                //writer.WriteElementString("FechaBaja", t105.lista_PM[i].fechaBaja.ToString("yyyy-MM-dd"));

                //                writer.WriteStartElement("Aparatos");
                //                {
                //                    for (int j = 0; j < t105.lista_PM[i].lista_aparatos.Count; j++)
                //                    {
                //                        writer.WriteStartElement("Aparato");
                //                        {
                //                            writer.WriteStartElement("ModeloAparato");
                //                            {
                //                                writer.WriteElementString("TipoAparato", t105.lista_PM[i].lista_aparatos[j].tipoAparato);
                //                                writer.WriteElementString("MarcaAparato", t105.lista_PM[i].lista_aparatos[j].marcaAparato);
                //                                writer.WriteElementString("ModeloMarca", t105.lista_PM[i].lista_aparatos[j].modeloMarca);
                //                            }
                //                            writer.WriteEndElement(); // ModeloAparato

                //                            writer.WriteElementString("TipoMovimiento", t105.lista_PM[i].lista_aparatos[j].tipoMovimiento);
                //                            writer.WriteElementString("TipoEquipoMedida", t105.lista_PM[i].lista_aparatos[j].tipoEquipoMedida);
                //                            writer.WriteElementString("TipoPropiedadAparato", t105.lista_PM[i].lista_aparatos[j].tipoPropiedadAparato);
                //                            writer.WriteElementString("TipoDHEdM", t105.lista_PM[i].lista_aparatos[j].tipoDHEdM);
                //                            writer.WriteElementString("ModoMedidaPotencia", t105.lista_PM[i].lista_aparatos[j].modoMedidaPotencia);
                //                            writer.WriteElementString("CodPrecinto", t105.lista_PM[i].lista_aparatos[j].codPrecinto);

                //                            writer.WriteStartElement("DatosAparato");
                //                            {
                //                                writer.WriteElementString("PeriodoFabricacion", t105.lista_PM[i].lista_aparatos[j].periodoFabricacion);
                //                                writer.WriteElementString("NumeroSerie", t105.lista_PM[i].lista_aparatos[j].numeroSerie);
                //                                writer.WriteElementString("FuncionAparato", t105.lista_PM[i].lista_aparatos[j].funcionAparato);
                //                                writer.WriteElementString("NumIntegradores", t105.lista_PM[i].lista_aparatos[j].numIntegradores);
                //                                writer.WriteElementString("ConstanteEnergia", t105.lista_PM[i].lista_aparatos[j].constanteEnergia);
                //                                writer.WriteElementString("ConstanteMaximetro", t105.lista_PM[i].lista_aparatos[j].constanteMaximetro);
                //                                writer.WriteElementString("RuedasEnteras", t105.lista_PM[i].lista_aparatos[j].ruedasEnteras);
                //                                writer.WriteElementString("RuedasDecimales", t105.lista_PM[i].lista_aparatos[j].ruedasDecimales);

                //                            }
                //                            writer.WriteEndElement(); // DatosAparato

                //                            if (t105.lista_PM[i].lista_aparatos[j].listaMedidas.Count > 0)
                //                            {
                //                                writer.WriteStartElement("Medidas");
                //                                {
                //                                    foreach (EndesaEntity.contratacion.xxi.XML_Medidas p in t105.lista_PM[i].lista_aparatos[j].listaMedidas)
                //                                    {
                //                                        writer.WriteStartElement("Medida");
                //                                        writer.WriteElementString("TipoDHEdM", p.tipoDHEdM);
                //                                        writer.WriteElementString("Periodo", p.periodo);
                //                                        writer.WriteElementString("MagnitudMedida", p.magnitudMedida);
                //                                        writer.WriteElementString("Procedencia", p.procedencia);
                //                                        writer.WriteElementString("UltimaLecturaFirme", p.ultimaLecturaFirme);
                //                                        writer.WriteElementString("FechaLecturaFirme", p.fechaLecturaFirme.ToString("yyyy-MM-dd"));
                //                                        writer.WriteEndElement(); // Medida
                //                                    }
                //                                }
                //                                writer.WriteEndElement(); // Medidas

                //                            }
                //                        }
                //                        writer.WriteEndElement(); // Aparato

                //                    }
                //                }
                //                writer.WriteEndElement(); // Aparatos

                //            }
                //            writer.WriteEndElement(); // PuntoDeMedida
                //        }
                //        writer.WriteEndElement(); // PuntosDeMedida

                //    }
                //    writer.WriteEndElement(); // ActivacionCambiodeComercializadorSinCambios                  

                //}


                //writer.WriteEndElement(); //VueltasaCUR
                //writer.WriteEndDocument();

                //writer.Close();

                #endregion
                return false;
            }catch(Exception ex)
            {
                return true;
            }

        }

        public bool CreaXML_t101_A302_A305_v2(FileInfo file, EndesaEntity.contratacion.xxi.XML_Datos t101,
            EndesaEntity.contratacion.xxi.XML_Datos t105)
        {

            string rr = "*****************";

            try
            {


                #region RegistrodePuntodeSuministro
                if (t101.pais == "")
                    t101.pais = t101.cups.Substring(0, 2);
                if (t101.provincia == "")
                    t101.provincia = "NN";
                if (t101.municipio == "")
                    t101.municipio = "NNNNN";
                if (t101.poblacion == "")
                    t101.poblacion = "NNNNNNNNN";
                if (t101.descripcionPoblacion == "")
                    t101.descripcionPoblacion = rr;
                if (t101.tipoVia == "")
                    t101.tipoVia = "**";
                if (t101.codPostal == "")
                    t101.codPostal = "NNNNN";
                if (t101.calle == "")
                    t101.calle = rr;
                if (t101.numeroFinca == "")
                    t101.numeroFinca = rr;
                if (t101.piso == "")
                    t101.piso = rr;

                if (t105.indicativoDeInterrumpibilidad == "")
                    t105.indicativoDeInterrumpibilidad = "N";

                if (t101.tipoIdentificador == "")
                    t101.tipoIdentificador = "NI";
                if (t101.identificador == "")
                    t101.identificador = rr;
                if (t101.tipoPersona == "")
                    t101.tipoPersona = "J";
                //if (t101.cliente_NombreDePila == "")
                //    t101.cliente_NombreDePila = rr;
                //if (t101.cliente_PrimerApellido == "")
                //    t101.cliente_PrimerApellido = rr;

                if (t101.contacto_PersonaDeContacto == "")
                    t101.contacto_PersonaDeContacto = "****************************";

                if (t101.contacto_PrefijoPais == "")
                    t101.contacto_PrefijoPais = "NNNN";

                if (t101.contacto_Numero == "")
                    t101.contacto_Numero = "NNNNNNNNNNNN";

                if (t101.numero == "")
                    t101.numero = "NNNNNNNNNNNN";


                if (t101.descripcionPoblacion == "")
                    t101.descripcionPoblacion = rr;

                if (t101.cnae == null)
                    t101.cnae = "NNNN";

                #region "DireccionPS
                {
                    if (t101.direccionPS_Pais == "")
                        t101.direccionPS_Pais = t101.cups.Substring(0, 2);

                    if (t101.direccionPS_Provincia == "")
                        t101.direccionPS_Provincia = "NN";

                    if (t101.direccionPS_Municipio == "")
                        t101.direccionPS_Municipio = "NNNNN";

                    if (t101.direccionPS_Poblacion == "")
                        t101.direccionPS_Poblacion = "NNNNNNNNN";

                    if (t101.direccionPS_CodPostal == "")
                        t101.direccionPS_CodPostal = "NNNNN";

                    if (t101.direccionPS_Calle == "")
                        t101.direccionPS_Calle = rr;

                    if (t101.direccionPS_NumeroFinca == "")
                        t101.direccionPS_NumeroFinca = "SN";

                    if (t101.direccionPS_AclaradorFinca == "" || t101.direccionPS_AclaradorFinca == null)
                        t101.direccionPS_AclaradorFinca = rr;


                }
                #endregion

                #region Direccion Cliente
                {
                    if (t101.paisCliente == "")
                        t101.paisCliente = t101.cups.Substring(0, 2);

                    if (t101.provinciaCliente == "")
                        t101.provinciaCliente = "NN";

                    if (t101.municipioCliente == "")
                        t101.municipioCliente = "NNNNN";

                    if (t101.poblacionCliente == "")
                        t101.poblacionCliente = "NNNNNNNNN";


                    if (t101.codPostalCliente == "")
                        t101.codPostalCliente = "NNNNN";

                    if (t101.calleCliente == "")
                        t101.calleCliente = rr;

                    if (t101.numeroFincaCliente == "")
                        t101.numeroFincaCliente = "SN";

                    if (t101.aclaradorFinca == "")
                        t101.aclaradorFinca = rr;

                    if (t101.prefijoPais == "" || t101.prefijoPais == null)
                        t101.prefijoPais = "34";

                    if (t101.numero == "" || t101.numero == null)
                        t101.numero = "000000000";




                }
                #endregion


                if (t101.razonSocial == "")
                    t101.razonSocial = rr;



                if (t101.indicativoDeDireccionExterna == "S")
                {
                    if (t101.linea1DeLaDireccionExterna == "")
                        t101.linea1DeLaDireccionExterna = rr;
                    if (t101.linea2DeLaDireccionExterna == "")
                        t101.linea2DeLaDireccionExterna = rr;
                    if (t101.linea3DeLaDireccionExterna == "")
                        t101.linea3DeLaDireccionExterna = rr;
                    if (t101.linea4DeLaDireccionExterna == "")
                        t101.linea4DeLaDireccionExterna = rr;
                }



                #endregion

                EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeT101 _t101 =
                   new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeT101();

                 _t101.Cabecera.CodigoREEEmpresaEmisora = t101.codigoREEEmpresaEmisora;
                 _t101.Cabecera.CodigoREEEmpresaDestino = t101.codigoREEEmpresaDestino;
                 _t101.Cabecera.CodigoDelProceso = "T1";
                 _t101.Cabecera.CodigoDePaso = "01";
                 _t101.Cabecera.CodigoDeSolicitud = t101.codigoDeSolicitud;
                 _t101.Cabecera.SecuencialDeSolicitud = t101.secuencialDeSolicitud;
                 _t101.Cabecera.FechaSolicitud = t101.fechaSolicitud.ToString("yyyy-MM-ddTHH:mm:ss");
                 _t101.Cabecera.CUPS = t101.cups;
 
                 _t101.SolicitudTraspasoCOR.DatosSolicitud.motivoTraspaso = "01";
                 _t101.SolicitudTraspasoCOR.DatosSolicitud.fechaPrevistaAccion = DateTime.Now.ToString("yyyy-MM-dd");
                 _t101.SolicitudTraspasoCOR.DatosSolicitud.cnae = t101.cnae;
                 _t101.SolicitudTraspasoCOR.DatosSolicitud.indEsencial = "N";
                 // _t101.SolicitudTraspasoCOR.DatosSolicitud.suspBajaImpagoEnCurso = "N"; // Comentar a Victor

                 if (t101.fechaFinalizacion > DateTime.MinValue)
                    _t101.SolicitudTraspasoCOR.Contrato.FechaFinalizacion = t101.fechaFinalizacion.ToString("yyyy-MM-dd");

                _t101.SolicitudTraspasoCOR.Contrato.TipoAutoconsumo = t105.tipoAutoconsumo;
                _t101.SolicitudTraspasoCOR.Contrato.TipoContratoATR = t105.tipoContratoATR;

                _t101.SolicitudTraspasoCOR.Contrato.CondicionesContractuales.TarifaATR = t101.tarifaATR;


                EndesaEntity.cnmc.V21_2019_12_17.PotenciasContratadas potencias_contratadas =
                   new EndesaEntity.cnmc.V21_2019_12_17.PotenciasContratadas();

                EndesaEntity.cnmc.V21_2019_12_17.Potencia potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "1";
                potencia.potencia = t105.potenciaPeriodo[1].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "2";
                potencia.potencia = t105.potenciaPeriodo[2].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "3";
                potencia.potencia = t105.potenciaPeriodo[3].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "4";
                potencia.potencia = t105.potenciaPeriodo[4].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "5";
                potencia.potencia = t105.potenciaPeriodo[5].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                potencia = new EndesaEntity.cnmc.V21_2019_12_17.Potencia();
                potencia.periodo = "6";
                potencia.potencia = t105.potenciaPeriodo[6].ToString();
                potencias_contratadas.Potencia.Add(potencia);

                _t101.SolicitudTraspasoCOR.Contrato.CondicionesContractuales.PotenciasContratadas = potencias_contratadas;              
                

                if (t101.periodicidadFacturacion != "")
                    _t101.SolicitudTraspasoCOR.Contrato.CondicionesContractuales.PeriodicidadFacturacion = t101.periodicidadFacturacion;



                _t101.SolicitudTraspasoCOR.Contrato.Contacto.PersonaDeContacto = t101.contacto_PersonaDeContacto;
                Telefono telefono = new Telefono();
                telefono.Numero = t101.contacto_Numero;
                telefono.PrefijoPais = t101.contacto_PrefijoPais;

                _t101.SolicitudTraspasoCOR.Contrato.Contacto.Telefono = telefono;
                _t101.SolicitudTraspasoCOR.Contrato.Contacto.Telefono.Numero  = t101.contacto_Numero;

                _t101.SolicitudTraspasoCOR.Cliente.IdCliente.TipoIdentificador = t101.tipoIdentificador;
                _t101.SolicitudTraspasoCOR.Cliente.IdCliente.Identificador = t101.identificador;
                _t101.SolicitudTraspasoCOR.Cliente.IdCliente.TipoPersona = t101.tipoPersona;
                _t101.SolicitudTraspasoCOR.Cliente.Nombre.RazonSocial = t101.razonSocial;

                telefono = new Telefono();
                telefono.Numero = t101.numero;
                telefono.PrefijoPais = t101.prefijoPais;

                _t101.SolicitudTraspasoCOR.Cliente.Telefono = telefono;
                
                //_t101.SolicitudTraspasoCOR.Cliente.IndicadorTipoDireccion = t101.indicadorTipoDireccion;
                _t101.SolicitudTraspasoCOR.Cliente.IndicadorTipoDireccion = "S";

                _t101.SolicitudTraspasoCOR.DireccionPS.Pais = t101.direccionPS_Pais;
                _t101.SolicitudTraspasoCOR.DireccionPS.Provincia = t101.direccionPS_Provincia;
                _t101.SolicitudTraspasoCOR.DireccionPS.Municipio = t101.direccionPS_Municipio;
                _t101.SolicitudTraspasoCOR.DireccionPS.TipoVia = t101.direccionPS_TipoVia;
                _t101.SolicitudTraspasoCOR.DireccionPS.CodPostal = t101.direccionPS_CodPostal;
                _t101.SolicitudTraspasoCOR.DireccionPS.Calle = t101.direccionPS_Calle;
                _t101.SolicitudTraspasoCOR.DireccionPS.NumeroFinca = t101.direccionPS_NumeroFinca;

                XmlWriterSettings settings = new XmlWriterSettings();
                settings.OmitXmlDeclaration = true;
                settings.Indent = true;
                XmlWriter writer = XmlWriter.Create(file.FullName, settings);

                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add("", @"http://localhost/elegibilidad");

                XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeT101));
                serializer.Serialize(writer, _t101, ns);
                writer.Close();

                #region Comentario
                
               

                #endregion
                return false;
            }
            catch (Exception ex)
            {
                return true;
            }

        }
        public void CreaXML_T102(FileInfo file, EndesaEntity.contratacion.xxi.XML_Datos t101)
        {

            //string url = "http://xmlns.endesa.com/wsdl/FormatoCUR/";
            string url = "http://localhost/elegibilidad";

            XmlTextWriter writer = new XmlTextWriter(file.FullName, Encoding.UTF8);


            writer.WriteStartDocument();
            writer.WriteStartElement("MensajeRechazoTraspasoCOR", url);
            {
                writer.WriteStartElement("Cabecera");
                {
                    writer.WriteElementString("CodigoREEEmpresaEmisora", "0636");
                    writer.WriteElementString("CodigoREEEmpresaDestino", t101.codigoREEEmpresaDestino);
                    writer.WriteElementString("CodigoDelProceso", "T1");
                    writer.WriteElementString("CodigoDePaso", "02");
                    writer.WriteElementString("CodigoDeSolicitud", t101.codigoDeSolicitud);
                    writer.WriteElementString("SecuencialDeSolicitud", "00");
                    writer.WriteElementString("FechaSolicitud", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"));
                    writer.WriteElementString("CUPS", t101.cups);
                }
                writer.WriteEndElement(); // Cabecera

                writer.WriteStartElement("Rechazos");
                {
                    writer.WriteElementString("FechaRechazo", DateTime.Now.ToString("yyyy-MM-dd"));
                    writer.WriteStartElement("Rechazo");
                    {
                        writer.WriteElementString("Secuencial", "00");
                        writer.WriteElementString("CodigoMotivo", "E4");
                        writer.WriteElementString("Comentarios", "Impago Previo");
                    }
                    writer.WriteEndElement(); // Rechazo
                }
                writer.WriteEndElement(); // Rechazos

            }

            writer.WriteEndElement(); //VueltasaCUR
            writer.WriteEndDocument();
            writer.Close();
            
        }


        public void CreaXML_A301(FileInfo file, EndesaEntity.contratacion.xxi.Eventuales datos)
        {
            string url = "http://localhost/elegibilidad";

            XmlTextWriter writer = new XmlTextWriter(file.FullName, Encoding.UTF8);

            writer.WriteStartDocument();


            writer.WriteStartElement("MensajeAlta", url);
            {
                writer.WriteStartElement("Cabecera");
                {
                    writer.WriteElementString("CodigoREEEmpresaEmisora", "0636");
                    writer.WriteElementString("CodigoREEEmpresaDestino", datos.codigoREEEmpresaDestino);
                    writer.WriteElementString("CodigoDelProceso", "A3");
                    writer.WriteElementString("CodigoDePaso", "01");
                    writer.WriteElementString("SecuencialDeSolicitud", "00");
                    writer.WriteElementString("FechaSolicitud", datos.fechaSolicitud.ToString("yyyy-MM-ddTHH:mm:ss"));
                    writer.WriteElementString("CUPS", datos.cups);
                }
                writer.WriteEndElement(); // Cabecera

                writer.WriteStartElement("Alta");
                {
                    writer.WriteStartElement("DatosSolicitud");
                    {
                        writer.WriteElementString("CNAE", datos.cnae.ToString().PadLeft(4, '0'));
                        writer.WriteElementString("IndActivacion", datos.indActivacion);
                        writer.WriteElementString("FechaPrevistaAccion", datos.fechaPrevistaAccion.ToString("yyyy-MM-dd"));
                        writer.WriteElementString("SolicitudTension", "N");
                    }
                    writer.WriteEndElement(); // DatosSolicitud

                    writer.WriteStartElement("Contrato");
                    {
                        writer.WriteElementString("FechaFinalizacion", datos.fecha_de_baja.ToString("yyyy-MM-dd"));
                        writer.WriteElementString("TipoAutoconsumo","00");
                        writer.WriteElementString("TipoContratoATR", datos.tipoContratoATR);
                        writer.WriteElementString("SolicitudTension", "N");

                        writer.WriteStartElement("CondicionesContractuales");
                        {
                            writer.WriteElementString("TarifaATR", datos.tarifaATR);

                            writer.WriteStartElement("PotenciasContratadas");
                            {
                                for (int i = 1; i < 7; i++)
                                {
                                    if (datos.potenciaPeriodo[i] > 0)
                                    {
                                        writer.WriteStartElement("Potencia");
                                        writer.WriteAttributeString("Periodo", i.ToString());
                                        writer.WriteString(datos.potenciaPeriodo[i].ToString());
                                        writer.WriteEndElement();                                        
                                    }                                    
                                }
                            }
                            writer.WriteEndElement(); // PotenciasContratadas                            
                        }
                        writer.WriteEndElement(); // CondicionesContractuales

                        writer.WriteStartElement("Contacto");
                        {
                            writer.WriteStartElement("PersonaDeContacto");                            
                            writer.WriteString(datos.contacto_PersonaDeContacto);
                            writer.WriteEndElement();

                            writer.WriteStartElement("Telefono");
                            {
                                writer.WriteElementString("PrefijoPais", "0034");
                                writer.WriteElementString("Numero", datos.contacto_Numero);
                            }
                            writer.WriteEndElement(); // PotenciasContratadas
                            
                        }
                        writer.WriteEndElement(); // Contacto  

                    }
                    writer.WriteEndElement(); // Contrato

                    writer.WriteStartElement("Cliente");
                    {
                        writer.WriteStartElement("IdCliente");
                        {
                            writer.WriteElementString("TipoIdentificador", datos.tipoIdentificador);
                            writer.WriteElementString("Identificador", datos.identificador);
                        }
                        writer.WriteEndElement(); // IdCliente 

                        writer.WriteStartElement("Nombre");
                        {
                            writer.WriteElementString("RazonSocial", datos.razonSocial);                            
                        }
                        writer.WriteEndElement(); // Nombre 

                        writer.WriteStartElement("Direccion");
                        {
                            writer.WriteElementString("Pais", "ESPANA");
                            writer.WriteElementString("Provincia", datos.provinciaCliente);
                            writer.WriteElementString("Municipio", datos.municipioCliente);
                            writer.WriteElementString("CodPostal", datos.codPostalCliente);

                            writer.WriteStartElement("Via");
                            {
                                writer.WriteElementString("TipoVia", "ESPANA");
                                writer.WriteElementString("Calle", datos.provinciaCliente);
                                writer.WriteElementString("NumeroFinca", datos.municipioCliente);
                                writer.WriteElementString("DuplicadorFinca", "SIN");
                                writer.WriteElementString("TipoAclaradorFinca", "BI");
                                writer.WriteElementString("AclaradorFinca", "SIN");
                            }
                            writer.WriteEndElement(); // Via

                        }
                        writer.WriteEndElement(); // Direccion 

                    }
                    writer.WriteEndElement(); // Cliente 

                    if(datos.observaciones != null)
                        writer.WriteElementString("Comentarios", datos.observaciones);
                    else
                        writer.WriteElementString("Comentarios", "SIN COMENTARIOS");


                    writer.WriteStartElement("RegistrosDocumento");
                    {
                        writer.WriteStartElement("RegistroDoc");
                        {
                            writer.WriteElementString("TipoDocAportado", datos.tipoDocAportado);
                            writer.WriteElementString("DireccionUrl", datos.direccionURL);
                        }
                        writer.WriteEndElement(); // RegistroDoc 
                    }
                    writer.WriteEndElement(); // RegistrosDocumento 

                }
                writer.WriteEndElement(); // Alta
            }
            writer.WriteEndElement(); // MensajeAlta

        }

        //public void GuardadoBBDD_T101(Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic)
        //{
        //    MySQLDB db;
        //    MySqlCommand command;
        //    string strSql = "";
        //    bool firstOnly = true;
        //    int num_reg = 0;

        //    foreach (KeyValuePair<string,EndesaEntity.contratacion.xxi.XML_Datos> p in dic)
        //    {
        //        EndesaEntity.contratacion.xxi.XML_Datos xml = p.Value;

        //        if (firstOnly)
        //        {
        //            strSql = "replace into cont_xml_t101 (CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso, CodigoDePaso, CodigoDeSolicitud," 
        //                + " SecuencialDeSolicitud, FechaSolicitud, CUPS, MotivoTraspaso, FechaPrevistaAccion, CNAE, IndEsencial," 
        //                + " TipoAutoconsumo, TipoContratoATR, TarifaATR, PotenciaPeriodo1, PotenciaPeriodo2, PotenciaPeriodo3," 
        //                + " PotenciaPeriodo4, PotenciaPeriodo5, PotenciaPeriodo6, ModoControlPotencia, PeriodicidadFacturacion," 
        //                + " Contacto_PersonaDeContacto, Contacto_PrefijoPais, Contacto_Numero, Contacto_CorreoElectronico," 
        //                + " Cliente_TipoIdentificador, Cliente_Identificador, Cliente_TipoPersona, Cliente_NombreDePila," 
        //                + " Cliente_PrimerApellido, Cliente_SegundoApellido, Cliente_PrefijoPais, Cliente_Numero, Cliente_CorreoElectronico," 
        //                + " Cliente_IndicadorTipoDireccion, Cliente_Pais, Cliente_Provincia, Cliente_Municipio, Cliente_Poblacion," 
        //                + " Cliente_CodPostal, Cliente_TipoVia, Cliente_Calle, Cliente_NumeroFinca, DireccionPS_Pais," 
        //                + " DireccionPS_Provincia, DireccionPS_Municipio, DireccionPS_Poblacion, DireccionPS_CodPostal," 
        //                + " DireccionPS_Calle, DireccionPS_NumeroFinca,"
        //                + " created_by, fichero) values ";
        //            firstOnly = false;
        //        }

        //        num_reg++;

        //        #region Campos

        //        if (xml.codigoREEEmpresaEmisora != null)
        //            strSql += "('" + xml.codigoREEEmpresaEmisora + "'";
        //        else
        //            strSql += "(null";

        //        if (xml.codigoREEEmpresaDestino != null)
        //            strSql += ", '" + xml.codigoREEEmpresaDestino + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codigoDelProceso != null)
        //            strSql += ", '" + xml.codigoDelProceso + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codigoDePaso != null)
        //            strSql += ", '" + xml.codigoDePaso + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.codigoDeSolicitud != null)
        //            strSql += ", '" + xml.codigoDeSolicitud + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.secuencialDeSolicitud != null)
        //            strSql += ", '" + xml.secuencialDeSolicitud + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.fechaSolicitud != null)
        //            strSql += ", '" + xml.fechaSolicitud.ToString("yyyy-MM-dd HH:mm:ss") + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cups != null)
        //            strSql += ", '" + xml.cups + "'";
        //        else
        //            strSql += ", null";

        //        if(xml.motivoTraspaso != null)
        //            strSql += ", '" + xml.motivoTraspaso + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.fechaPrevistaAccion != null)
        //            strSql += ", '" + xml.fechaPrevistaAccion.ToString("yyyy-MM-dd") + "'";
        //        else
        //            strSql += ", null";               


        //        if (xml.cnae != 0)
        //            strSql += ", " + xml.cnae;
        //        else
        //            strSql += ", null";

        //        if (xml.indEsencial != null)
        //            strSql += ", '" + xml.indEsencial + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipoAutoconsumo != null)
        //            strSql += ", '" + xml.tipoAutoconsumo + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tipoContratoATR != null)
        //            strSql += ", '" + xml.tipoContratoATR + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.tarifaATR != null)
        //            strSql += ", '" + xml.tarifaATR + "'";
        //        else
        //            strSql += ", null";

        //        for (int i = 1; i < xml.potenciaPeriodo.Count(); i++)
        //            if (xml.potenciaPeriodo[i] != 0)
        //                strSql += ", " + xml.potenciaPeriodo[i];
        //            else
        //                strSql += ", null";

        //        if (xml.modoControlPotencia != null)
        //            strSql += ", '" + xml.modoControlPotencia + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.periodicidadFacturacion != null)
        //            strSql += ", '" + xml.periodicidadFacturacion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.contacto_PersonaDeContacto != null)
        //            strSql += ", '" + xml.contacto_PersonaDeContacto + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.contacto_PrefijoPais != null)
        //            strSql += ", '" + xml.contacto_PrefijoPais + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.contacto_Numero != null)
        //            strSql += ", '" + xml.contacto_Numero + "'"; 
        //        else
        //            strSql += ", null";


        //        if (xml.contacto_CorreoElectronico != null)
        //            strSql += ", '" + xml.contacto_CorreoElectronico + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.tipoIdentificador != null)
        //            strSql += ", '" + xml.tipoIdentificador + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.identificador != null)
        //            strSql += ", '" + xml.identificador + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.tipoPersona != null)
        //            strSql += ", '" + xml.tipoPersona + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.razonSocial != null)
        //            strSql += ", '" + xml.razonSocial + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_PrimerApellido != null)
        //            strSql += ", '" + xml.cliente_PrimerApellido + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_SegundoApellido != null)
        //            strSql += ", '" + xml.cliente_SegundoApellido + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.cliente_PrefijoPais != null)
        //            strSql += ", '" + xml.cliente_PrefijoPais + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_Numero != null)
        //            strSql += ", '" + xml.cliente_Numero + "'";
        //        else
        //            strSql += ", null";


        //        if (xml.cliente_CorreoElectronico != null)
        //            strSql += ", '" + xml.cliente_CorreoElectronico + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_IndicadorTipoDireccion != null)
        //            strSql += ", '" + xml.cliente_IndicadorTipoDireccion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_Pais != null)
        //            strSql += ", '" + xml.cliente_Pais + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_Provincia != null)
        //            strSql += ", '" + xml.cliente_Provincia + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_Municipio != null)
        //            strSql += ", '" + xml.cliente_Municipio + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_Poblacion != null)
        //            strSql += ", '" + xml.cliente_Poblacion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_CodPostal != null)
        //            strSql += ", '" + xml.cliente_CodPostal + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_TipoVia != null)
        //            strSql += ", '" + xml.cliente_TipoVia + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_Calle != null)
        //            strSql += ", '" + xml.cliente_Calle + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.cliente_NumeroFinca != null)
        //            strSql += ", '" + xml.cliente_NumeroFinca + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.direccionPS_Pais != null)
        //            strSql += ", '" + xml.direccionPS_Pais + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.direccionPS_Provincia != null)
        //            strSql += ", '" + xml.direccionPS_Provincia + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.direccionPS_Municipio != null)
        //            strSql += ", '" + xml.direccionPS_Municipio + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.direccionPS_Poblacion != null)
        //            strSql += ", '" + xml.direccionPS_Poblacion + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.direccionPS_CodPostal != null)
        //            strSql += ", '" + xml.direccionPS_CodPostal + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.direccionPS_Calle != null)
        //            strSql += ", '" + xml.direccionPS_Calle + "'";
        //        else
        //            strSql += ", null";

        //        if (xml.direccionPS_NumeroFinca != null)
        //            strSql += ", '" + xml.direccionPS_NumeroFinca + "'";
        //        else
        //            strSql += ", null";


        //        strSql += ", '" + System.Environment.UserName + "'" + ", '" + xml.fichero + "'),";
        //        #endregion

        //        if (num_reg > 250)
        //        {
        //            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
        //            command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
        //            command.ExecuteNonQuery();
        //            db.CloseConnection();
        //            num_reg = 0;
        //            strSql = "";
        //            firstOnly = true;
        //        }

        //    }

        //    if (num_reg > 0)
        //    {
        //        db = new EndesaBusiness.servidores.MySQLDB(EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
        //        command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
        //        command.ExecuteNonQuery();
        //        db.CloseConnection();
        //        num_reg = 0;
        //        strSql = "";
        //    }


        //}

        private void ValidationCallBack(object sender, ValidationEventArgs e)
        {
            XmlSeverityType type = XmlSeverityType.Warning;
            if (Enum.TryParse<XmlSeverityType>("Error", out type))
            {
                if (type == XmlSeverityType.Error) throw new Exception(e.Message);
            }
        }

        // SOLICITADO POR JOSE MANUEL GUERRERO - GUS: creamos esta función estática para un caso especifico
        // ATENCION, FUNCIONES TEMPORALES Y PARA UN CASO ESPECIFICO, PARAMETROS HARDCODEADO, SE DEJAN POR SI HACE FALTA USO POSTERIOR
        // Tenemos que extraer los CUPS que vienen en ficheros XML comprimidos en ZIP, y los valores se
        // encuentran en los nodos con nombre CPE o NR_CPE
        // Los insertamos en la tabla cont.rpe_masivo
        public static void SaveDBNodeValue(string pathName, string nodeName)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            db = new MySQLDB(MySQLDB.Esquemas.CON);

            FileInfo file_xml, file_zip;
            string cod_ini = "";
            string cod_fin = "";
            string valor = ""; 
            //string fileName = "";
            string pathTemp = "C:\\Temp\\rpe_masivo";

            //Obtenemos listado de los zip a procesar
            string[] zipFiles = System.IO.Directory.GetFiles(pathName, "*.zip");

            //Para cada fichero zip
            foreach (string fileName in zipFiles)
            {
                file_zip = new FileInfo(fileName);
                //Extraemos zip a carpeta temporal C:\Temp\rpe_masivo
                Console.WriteLine("Descomprimiendo zip " + fileName + " en " + pathTemp);
                EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();
                zip.DescomprimirArchivoZip(fileName, pathTemp);

                //Para cada XML extraido en ruta temporal pathTemp
                string[] xmlFiles = System.IO.Directory.GetFiles(pathTemp, "*.xml");
                Console.WriteLine("Procesando ficheros XML");
                foreach (string xmlfilename in xmlFiles) 
                {
                    
                    

                    Console.WriteLine("Extrayendo CUPS de fichero " + xmlfilename);
                    file_xml = new FileInfo(xmlfilename);
                    XmlTextReader r = new XmlTextReader(xmlfilename);

                    while (r.Read())
                    {
                        switch (r.NodeType)
                        {
                            case XmlNodeType.Element:
                                cod_ini = r.Name;
                                valor = "";
                                break;
                            case XmlNodeType.Text:
                                valor = r.Value;
                                break;
                            case XmlNodeType.EndElement:
                                cod_fin = r.Name;
                                break;
                        }
                        if ( (valor != "") && (cod_ini == cod_fin) && (cod_fin == nodeName || cod_fin == "NR_" + nodeName) )
                        {
                            //Guardar CUPS en base de datos
                            Console.WriteLine(cod_fin + ": " + valor + " | Fichero: " + file_zip.Name);

                            strSql = "REPLACE into rpe_masivo SET cups='"+ valor + "', filename='"+ file_zip.Name + "';";
                            command = new MySqlCommand(strSql, db.con);
                            command.CommandTimeout = 3000;
                            command.ExecuteNonQuery();
                            
                        }
                    }
                    
                    r.Close();
                    db.CloseConnection();
                    //Borramos fichero XML procesado
                    file_xml.Delete();
                }
                //Movemos a procesados
               File.Move(fileName, pathName + "\\procesados\\" + file_zip.Name);
            }

        }
        // Idem función anterior pero guarda en fichero formato csv 
        // ATENCION, FUNCIONES TEMPORALES Y PARA UN CASO ESPECIFICO, PARAMETROS HARDCODEADO, SE DEJAN POR SI HACE FALTA USO POSTERIOR
        public static void SaveFileDBNodeValue(string pathName, string nodeName)
        {

            FileInfo file_xml, file_zip;
            string cod_ini = "";
            string cod_fin = "";
            string valor = "";
            //string fileName = "";
            string pathTemp = "C:\\Temp\\rpe_masivo";

            //Obtenemos listado de los zip a procesar
            string[] zipFiles = System.IO.Directory.GetFiles(pathName, "*.zip");

            //Para cada fichero zip
            foreach (string fileName in zipFiles)
            {
                file_zip = new FileInfo(fileName);
                Console.WriteLine(System.IO.Path.GetFileNameWithoutExtension(file_zip.Name));

                //Creamos fichero de texto pathName + "\\salida\\" + file_zip.Name
                StreamWriter sw = File.CreateText(pathName + "\\salida\\" + System.IO.Path.GetFileNameWithoutExtension(file_zip.Name) + "_CUPS.csv");
                sw.WriteLine("cups;filename;");

                //Extraemos zip a carpeta temporal C:\Temp\rpe_masivo
                Console.WriteLine("Descomprimiendo zip " + fileName + " en " + pathTemp);
                EndesaBusiness.utilidades.ZIP zip = new EndesaBusiness.utilidades.ZIP();
                zip.DescomprimirArchivoZip(fileName, pathTemp);

                //Para cada XML extraido en ruta temporal pathTemp
                string[] xmlFiles = System.IO.Directory.GetFiles(pathTemp, "*.xml");
                Console.WriteLine("Procesando ficheros XML");
                foreach (string xmlfilename in xmlFiles)
                {

                    Console.WriteLine("Extrayendo CUPS de fichero " + xmlfilename);
                    file_xml = new FileInfo(xmlfilename);
                    XmlTextReader r = new XmlTextReader(xmlfilename);

                    while (r.Read())
                    {
                        switch (r.NodeType)
                        {
                            case XmlNodeType.Element:
                                cod_ini = r.Name;
                                valor = "";
                                break;
                            case XmlNodeType.Text:
                                valor = r.Value;
                                break;
                            case XmlNodeType.EndElement:
                                cod_fin = r.Name;
                                break;
                        }
                        if ((valor != "") && (cod_ini == cod_fin) && (cod_fin == nodeName || cod_fin == "NR_" + nodeName))
                        {
                            //Guardar CUPS en fichero formato csv
                            Console.WriteLine(cod_fin + ": " + valor + " | Fichero: " + file_zip.Name);
                            sw.WriteLine(valor+";"+file_zip.Name+";");
                        }
                    }

                    r.Close();
                   
                    //Borramos fichero XML procesado
                    file_xml.Delete();
                }
                //Movemos a procesados
                File.Move(fileName, pathName + "\\procesados\\" + file_zip.Name);
                sw.Close();
            }

        }
    }

}


