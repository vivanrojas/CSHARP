using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace EndesaBusiness.contratacion
{
    public class CNMC_XML
    {
        public List<EndesaEntity.contratacion.XML_CNMC> list { get; set; }
        utilidades.Param p;
        public CNMC_XML()
        {
            p = new utilidades.Param("eer_param", servidores.MySQLDB.Esquemas.CON);
            list = Carga();
        }

        private List<EndesaEntity.contratacion.XML_CNMC> Carga()
        {
            List<EndesaEntity.contratacion.XML_CNMC> l = new List<EndesaEntity.contratacion.XML_CNMC>();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "SELECT CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso, CodigoDePaso," +
                    " CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud, CUPS, CNAE, IndActivacion," +
                    " FechaPrevistaAccion, TipoActivacionPrevista, FechaActivacionPrevista, PotenciaExtension," +
                    " PotenciaDeAcceso, PotenciaInstAT, IndicativoDeInterrumpibilidad, Pais, Provincia, Municipio," +
                    " Poblacion, DescripcionPoblacion, TipoVia, CodPostal, Calle, NumeroFinca, AclaradorFinca," +
                    " TipoIdentificador, Identificador, TipoPersona, RazonSocial, PrefijoPais, Numero, CorreoElectronico," +
                    " IndicadorTipoDireccion, IndicativoDeDireccionExterna, Linea1DeLaDireccionExterna," +
                    " Linea2DeLaDireccionExterna, Linea3DeLaDireccionExterna, Linea4DeLaDireccionExterna, PaisCliente," +
                    " ProvinciaCliente, MunicipioCliente, PoblacionCliente, DescripcionPoblacionCliente, CodPostalCliente," +
                    " CalleCliente, NumeroFincaCliente, PisoCliente, Idioma, FechaActivacion, CodContrato, TipoAutoconsumo," +
                    " TipoContratoATR, TarifaATR, PeriodicidadFacturacion, TipodeTelegestion, PotenciaPeriodo1," +
                    " PotenciaPeriodo2, PotenciaPeriodo3, PotenciaPeriodo4, PotenciaPeriodo5, PotenciaPeriodo6," +
                    " MarcaMedidaConPerdidas, TensionDelSuministro, fichero, created_by, last_update_date" +
                    " FROM admP.eer_xml_tmp";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    //EndesaEntity.contratacion.XML_CNMC c = new EndesaEntity.contratacion.XML_CNMC();
                    //c.nif = r["nif"].ToString();
                    //c.cliente = r["cliente"].ToString();
                    //c.cpe = r["cpe"].ToString();
                    //c.click = Convert.ToInt32(r["click"]);
                    //c.mercado = r["mercado"].ToString();
                    //c.operacion = r["operacion"].ToString();
                    //c.fecha_operacion = Convert.ToDateTime(r["fecha"]);
                    //c.producto = r["producto"].ToString();
                    //c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);
                    //c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);
                    //c.bl = Convert.ToDouble(r["bl"]);
                    //c.fee = Convert.ToDouble(r["fee"]);
                    //c.volumen = Convert.ToDouble(r["volumen"]);

                    //List<EndesaEntity.facturacion.ClicksPT> o;
                    //if (!d.TryGetValue(c.cpe, out o))
                    //{
                    //    o = new List<EndesaEntity.facturacion.ClicksPT>();
                    //    o.Add(c);
                    //    d.Add(c.cpe, o);
                    //}
                    //else
                    //    o.Add(c);

                }

                db.CloseConnection();
                return l;

            }
            catch (Exception e)
            {

                MessageBox.Show("Carga completada satisfactoriamente.",
                  "Clicks.Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
                return null;
            }
        }
        public void CargaXML()
        {
            int total_archivos = 0;
            List<EndesaEntity.contratacion.XML_CNMC> lista = new List<EndesaEntity.contratacion.XML_CNMC>();
            OpenFileDialog d = new OpenFileDialog();
            d.Title = p.GetValue("mensaje_ventana_xml", DateTime.Now, DateTime.Now);
            d.Filter = "XML files|*.xml";
            d.Multiselect = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in d.FileNames)
                {
                    total_archivos++;
                    lista.Add(TrataXML(fileName));
                }

                GuardadoBBDD(lista);

                MessageBox.Show("La importación ha concluido correctamente."
                       + System.Environment.NewLine
                       + System.Environment.NewLine
                       + "Se han procesado " + lista.Count.ToString("#.###") + " archivos de " + total_archivos.ToString("#.###"),
                      "Importación ficheros XML",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

            }
        }

        private EndesaEntity.contratacion.XML_CNMC TrataXML(string fileName)
        {
            string cod_ini = "";
            string cod_fin = "";
            string valor = "";
            bool dentroCliente = false;

            int potencia = 0;

            EndesaEntity.contratacion.XML_CNMC xml = new EndesaEntity.contratacion.XML_CNMC();
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
                        break;
                    case XmlNodeType.Text: //Display the text in each element.
                        valor = utilidades.FuncionesTexto.ArreglaAcentos(r.Value);
                        break;
                    case XmlNodeType.EndElement: //Display the end of the element.
                        cod_fin = r.Name;
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
                            xml.cnae = Convert.ToInt32(valor);
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
                            else
                                xml.pais = valor;
                            break;
                        case "Provincia":
                            if (xml.provincia != null)
                                xml.provinciaCliente = valor;
                            else
                                xml.provincia = valor;
                            break;
                        case "Municipio":
                            if (xml.municipio != null)
                                xml.municipioCliente = valor;
                            else
                                xml.municipio = valor;
                            break;
                        case "Poblacion":
                            if (xml.poblacion != null)
                                xml.poblacionCliente = valor;
                            else
                                xml.poblacion = valor;
                            break;
                        case "DescripcionPoblacion":
                            if (xml.descripcionPoblacion != null)
                                xml.descripcionPoblacionCliente = valor;
                            else
                                xml.descripcionPoblacion = valor;
                            break;
                        case "TipoVia":
                            if (xml.tipoVia != null)
                                xml.tipoViaCliente = valor;
                            else
                                xml.tipoVia = valor;
                            break;
                        case "CodPostal":
                            if (xml.codPostal != null)
                                xml.codPostalCliente = valor;
                            else
                                xml.codPostal = valor;
                            break;
                        case "Calle":
                            if (xml.calle != null)
                                xml.calleCliente = valor;
                            else
                                xml.calle = valor;
                            break;
                        case "NumeroFinca":
                            if (xml.numeroFinca != null)
                                xml.numeroFincaCliente = valor;
                            else
                                xml.numeroFinca = valor;
                            break;
                        case "AclaradorFinca":
                            xml.aclaradorFinca = valor;
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
                        case "PrefijoPais":
                            xml.prefijoPais = valor;
                            break;
                        case "Numero":
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
                        case "IndActivacion":
                            xml.indActivacion = valor;
                            break;
                        case "FechaPrevistaAccion":
                            xml.fechaPrevistaAccion = Convert.ToDateTime(valor);
                            break;
                        case "TipoActivacionPrevista":
                            xml.tipoActivacionPrevista = valor;
                            break;
                        case "FechaActivacionPrevista":
                            xml.fechaActivacionPrevista = Convert.ToDateTime(valor);
                            break;
                    }


                #endregion

            }

            return xml;


        }

        private void GuardadoBBDD(List<EndesaEntity.contratacion.XML_CNMC> lista)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            int num_reg = 0;

            foreach (EndesaEntity.contratacion.XML_CNMC xml in lista)
            {
                if (firstOnly)
                {
                    strSql = "replace into eer_xml_tmp (CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,"
                        + " CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud, CUPS, CNAE," 
                        + " IndActivacion, FechaPrevistaAccion, TipoActivacionPrevista, FechaActivacionPrevista, PotenciaExtension,"
                        + " PotenciaDeAcceso, PotenciaInstAT, IndicativoDeInterrumpibilidad, Pais, Provincia, Municipio, Poblacion,"
                        + " DescripcionPoblacion, TipoVia, CodPostal, Calle, NumeroFinca, AclaradorFinca, TipoIdentificador, Identificador,"
                        + " TipoPersona, RazonSocial, PrefijoPais, Numero, CorreoElectronico, IndicadorTipoDireccion, IndicativoDeDireccionExterna,"
                        + " Linea1DeLaDireccionExterna, Linea2DeLaDireccionExterna, Linea3DeLaDireccionExterna, Linea4DeLaDireccionExterna,"
                        + " PaisCliente, ProvinciaCliente, MunicipioCliente, PoblacionCliente, DescripcionPoblacionCliente, CodPostalCliente,"
                        + " CalleCliente, NumeroFincaCliente, PisoCliente,"
                        + " Idioma, FechaActivacion, CodContrato, TipoAutoconsumo, TipoContratoATR, TarifaATR, PeriodicidadFacturacion, TipodeTelegestion,"
                        + " PotenciaPeriodo1, PotenciaPeriodo2, PotenciaPeriodo3, PotenciaPeriodo4, PotenciaPeriodo5, PotenciaPeriodo6, "
                        + " MarcaMedidaConPerdidas, TensionDelSuministro,"
                        + " created_by, fichero) values ";
                    firstOnly = false;
                }

                num_reg++;

                #region Campos

                if (xml.codigoREEEmpresaEmisora != null)
                    strSql += "('" + xml.codigoREEEmpresaEmisora + "'";
                else
                    strSql += "(null";

                if (xml.codigoREEEmpresaDestino != null)
                    strSql += ", '" + xml.codigoREEEmpresaDestino + "'";
                else
                    strSql += ", null";

                if (xml.codigoDelProceso != null)
                    strSql += ", '" + xml.codigoDelProceso + "'";
                else
                    strSql += ", null";

                if (xml.codigoDePaso != null)
                    strSql += ", '" + xml.codigoDePaso + "'";
                else
                    strSql += ", null";

                if (xml.codigoDeSolicitud != null)
                    strSql += ", '" + xml.codigoDeSolicitud + "'";
                else
                    strSql += ", null";

                if (xml.secuencialDeSolicitud != null)
                    strSql += ", '" + xml.secuencialDeSolicitud + "'";
                else
                    strSql += ", null";

                if (xml.fechaSolicitud > DateTime.MinValue)
                    strSql += ", '" + xml.fechaSolicitud.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                else
                    strSql += ", null";

                if (xml.cups != null)
                    strSql += ", '" + xml.cups + "'";
                else
                    strSql += ", null";

                if (xml.cnae != 0)
                    strSql += ", " + xml.cnae;
                else
                    strSql += ", null";

                if(xml.indActivacion != null)              
                    strSql += ", '" + xml.indActivacion + "'";
                else
                    strSql += ", null";
             
                if (xml.fechaPrevistaAccion > DateTime.MinValue)
                    strSql += ", '" + xml.fechaPrevistaAccion.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.tipoActivacionPrevista != null)
                    strSql += ", '" + xml.tipoActivacionPrevista + "'";
                else
                    strSql += ", null";

                if (xml.fechaActivacionPrevista > DateTime.MinValue)
                    strSql += ", '" + xml.fechaActivacionPrevista.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";


                if (xml.potenciaExtension != 0)
                    strSql += ", " + xml.potenciaExtension;
                else
                    strSql += ", null";

                if (xml.potenciaDeAcceso != 0)
                    strSql += ", " + xml.potenciaDeAcceso;
                else
                    strSql += ", null";


                if (xml.potenciaInstAT != 0)
                    strSql += ", " + xml.potenciaInstAT;
                else
                    strSql += ", null";

                if (xml.indicativoDeInterrumpibilidad != null)
                    strSql += ", '" + xml.indicativoDeInterrumpibilidad + "'";
                else
                    strSql += ", null";

                if (xml.pais != null)
                    strSql += ", '" + xml.pais + "'";
                else
                    strSql += ", null";

                if (xml.provincia != null)
                    strSql += ", '" + xml.provincia + "'";
                else
                    strSql += ", null";

                if (xml.municipio != null)
                    strSql += ", '" + xml.municipio + "'";
                else
                    strSql += ", null";

                if (xml.poblacion != null)
                    strSql += ", '" + xml.poblacion + "'";
                else
                    strSql += ", null";

                if (xml.descripcionPoblacion != null)
                    strSql += ", '" + xml.descripcionPoblacion + "'";
                else
                    strSql += ", null";

                if (xml.tipoVia != null)
                    strSql += ", '" + xml.tipoVia + "'";
                else
                    strSql += ", null";

                if (xml.codPostal != null)
                    strSql += ", '" + xml.codPostal + "'";
                else
                    strSql += ", null";

                if (xml.calle != null)
                    strSql += ", '" + xml.calle + "'";
                else
                    strSql += ", null";

                if (xml.numeroFinca != null)
                    strSql += ", '" + xml.numeroFinca + "'";
                else
                    strSql += ", null";

                if (xml.aclaradorFinca != null)
                    strSql += ", '" + xml.aclaradorFinca + "'";
                else
                    strSql += ", null";

                if (xml.tipoIdentificador != null)
                    strSql += ", '" + xml.tipoIdentificador + "'";
                else
                    strSql += ", null";

                if (xml.identificador != null)
                    strSql += ", '" + xml.identificador + "'";
                else
                    strSql += ", null";

                if (xml.tipoPersona != null)
                    strSql += ", '" + xml.tipoPersona + "'";
                else
                    strSql += ", null";

                if (xml.razonSocial != null)
                    strSql += ", '" + xml.razonSocial + "'";
                else
                    strSql += ", null";

                if (xml.prefijoPais != null)
                    strSql += ", '" + xml.prefijoPais + "'";
                else
                    strSql += ", null";

                if (xml.numero != null)
                    strSql += ", '" + xml.numero + "'";
                else
                    strSql += ", null";

                if (xml.correoElectronico != null)
                    strSql += ", '" + xml.correoElectronico + "'";
                else
                    strSql += ", null";

                if (xml.indicadorTipoDireccion != null)
                    strSql += ", '" + xml.indicadorTipoDireccion + "'";
                else
                    strSql += ", null";

                if (xml.indicativoDeDireccionExterna != null)
                    strSql += ", '" + xml.indicativoDeDireccionExterna + "'";
                else
                    strSql += ", null";

                if (xml.linea1DeLaDireccionExterna != null)
                    strSql += ", '" + xml.linea1DeLaDireccionExterna + "'";
                else
                    strSql += ", null";

                if (xml.linea2DeLaDireccionExterna != null)
                    strSql += ", '" + xml.linea2DeLaDireccionExterna + "'";
                else
                    strSql += ", null";

                if (xml.linea3DeLaDireccionExterna != null)
                    strSql += ", '" + xml.linea3DeLaDireccionExterna + "'";
                else
                    strSql += ", null";

                if (xml.linea4DeLaDireccionExterna != null)
                    strSql += ", '" + xml.linea4DeLaDireccionExterna + "'";
                else
                    strSql += ", null";

                if (xml.paisCliente != null)
                    strSql += ", '" + xml.paisCliente + "'";
                else
                    strSql += ", null";

                if (xml.provinciaCliente != null)
                    strSql += ", '" + xml.provinciaCliente + "'";
                else
                    strSql += ", null";

                if (xml.municipioCliente != null)
                    strSql += ", '" + xml.municipioCliente + "'";
                else
                    strSql += ", null";

                if (xml.poblacionCliente != null)
                    strSql += ", '" + xml.poblacionCliente + "'";
                else
                    strSql += ", null";

                if (xml.descripcionPoblacionCliente != null)
                    strSql += ", '" + xml.descripcionPoblacionCliente + "'";
                else
                    strSql += ", null";

                if (xml.codPostalCliente != null)
                    strSql += ", '" + xml.codPostalCliente + "'";
                else
                    strSql += ", null";

                if (xml.calleCliente != null)
                    strSql += ", '" + xml.calleCliente + "'";
                else
                    strSql += ", null";

                if (xml.numeroFincaCliente != null)
                    strSql += ", '" + xml.numeroFincaCliente + "'";
                else
                    strSql += ", null";

                if (xml.pisoCliente != null)
                    strSql += ", '" + xml.pisoCliente + "'";
                else
                    strSql += ", null";

                if (xml.idioma != null)
                    strSql += ", '" + xml.idioma + "'";
                else
                    strSql += ", null";

                if (xml.fechaActivacion > DateTime.MinValue)
                    strSql += ", '" + xml.fechaActivacion.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ", null";

                if (xml.codContrato != null)
                    strSql += ", '" + xml.codContrato + "'";
                else
                    strSql += ", null";

                if (xml.tipoAutoconsumo != null)
                    strSql += ", '" + xml.tipoAutoconsumo + "'";
                else
                    strSql += ", null";

                if (xml.tipoContratoATR != null)
                    strSql += ", '" + xml.tipoContratoATR + "'";
                else
                    strSql += ", null";

                if (xml.tarifaATR != null)
                    strSql += ", '" + xml.tarifaATR + "'";
                else
                    strSql += ", null";

                if (xml.periodicidadFacturacion != null)
                    strSql += ", '" + xml.periodicidadFacturacion + "'";
                else
                    strSql += ", null";

                if (xml.tipodeTelegestion != null)
                    strSql += ", '" + xml.tipodeTelegestion + "'";
                else
                    strSql += ", null";

                for (int i = 1; i < xml.potenciaPeriodo.Count(); i++)
                    if (xml.potenciaPeriodo[i] != 0)
                        strSql += ", " + xml.potenciaPeriodo[i];
                    else
                        strSql += ", null";

                if (xml.marcaMedidaConPerdidas != null)
                    strSql += ", '" + xml.marcaMedidaConPerdidas + "'";
                else
                    strSql += ", null";

                if (xml.tensionDelSuministro != 0)
                    strSql += ", " + xml.tensionDelSuministro;
                else
                    strSql += ", null";

                strSql += ", '" + System.Environment.UserName + "'" + ", '" + xml.fichero + "'),";
                #endregion

                if (num_reg > 250)
                {
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    num_reg = 0;
                    strSql = "";
                    firstOnly = true;
                }

            }

            if (num_reg > 0)
            {
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                num_reg = 0;
                strSql = "";
            }


        }
    }
}
