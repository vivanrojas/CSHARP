using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class EEXXI
    {
        public Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_XML;

        SolicitudesCodigos sc;
        Inventario inventario;

        InventarioDetalleEstados inventarioDetalleEstados;
        contratacion.PS_AT_Funciones ps;
        List<EndesaEntity.contratacion.xxi.InformeCasos> informe;
        EndesaBusiness.utilidades.Param param;

        EndesaBusiness.global.Provincias provincias;
        EndesaBusiness.global.Municipios municipio;
              

        public EEXXI()
        {
            param = new EndesaBusiness.utilidades.Param("eexxi_param", MySQLDB.Esquemas.CON);
            dic_XML = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();

            provincias = new global.Provincias("eexxi_param_provincias", servidores.MySQLDB.Esquemas.CON);
            municipio = new global.Municipios();
        }


        public void ConsultasActualizaDatosSolicitudes()
        {
            //Actualiza datos de direcciones, NIF y Razon Social desde t101

            MySQLDB db;
            MySqlCommand command;            
            string strSql;

            strSql = "UPDATE eexxi_solicitudes s"
                + " INNER JOIN eexxi_solicitudes_t101 t101 ON"
                + " t101.CodigoDeSolicitud = s.CodigoDeSolicitud AND"
                + " t101.CUPS = s.CUPS"
                + " SET s.Pais = t101.PaisCliente,"
                + " s.Provincia = t101.ProvinciaCliente,"
                + " s.Municipio = t101.MunicipioCliente,"
                + " s.Calle = t101.CalleCliente,"
                + " s.CodPostal = t101.CodPostalCliente,"
                + " s.NumeroFinca = t101.NumeroFincaCliente,"
                + " s.Identificador = t101.Identificador," 
                + " s.RazonSocial = t101.RazonSocial,"
                + " s.Linea1DeLaDireccionExterna = t101.Linea1DeLaDireccionExterna,"
                + " s.Linea2DeLaDireccionExterna = t101.Linea2DeLaDireccionExterna,"
                + " s.Linea3DeLaDireccionExterna = t101.Linea3DeLaDireccionExterna,"
                + " s.Linea4DeLaDireccionExterna = t101.Linea4DeLaDireccionExterna"
                + " WHERE s.IndicadorTipoDireccion = 'S' AND"
                + " (s.Pais = '' AND s.Provincia = '' AND s.Municipio = '' AND s.Poblacion = '')"
                + " OR s.Linea1DeLaDireccionExterna = ''";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE eexxi_solicitudes_tmp s"
                + " INNER JOIN eexxi_solicitudes_t101 t101 ON"
                + " t101.CodigoDeSolicitud = s.CodigoDeSolicitud AND"
                + " t101.CUPS = s.CUPS"
                + " SET s.Pais = t101.PaisCliente,"
                + " s.Provincia = t101.ProvinciaCliente,"
                + " s.Municipio = t101.MunicipioCliente,"
                + " s.Calle = t101.CalleCliente,"
                + " s.CodPostal = t101.CodPostalCliente,"
                + " s.NumeroFinca = t101.NumeroFincaCliente,"
                + " s.Identificador = t101.Identificador,"
                + " s.RazonSocial = t101.RazonSocial,"
                + " s.Linea1DeLaDireccionExterna = t101.Linea1DeLaDireccionExterna,"
                + " s.Linea2DeLaDireccionExterna = t101.Linea2DeLaDireccionExterna,"
                + " s.Linea3DeLaDireccionExterna = t101.Linea3DeLaDireccionExterna,"
                + " s.Linea4DeLaDireccionExterna = t101.Linea4DeLaDireccionExterna"
                + " WHERE s.IndicadorTipoDireccion = 'S' AND"
                + " (s.Pais = '' AND s.Provincia = '' AND s.Municipio = '' AND s.Poblacion = '')"
                + " OR s.Linea1DeLaDireccionExterna = ''";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


            strSql = "UPDATE eexxi_solicitudes s"
                + " INNER JOIN eexxi_solicitudes_t101 t101 ON"
                + " t101.CodigoDeSolicitud = s.CodigoDeSolicitud AND"
                + " t101.CUPS = s.CUPS"
                + " SET s.PaisCliente = t101.PaisCliente,"
                + " s.ProvinciaCliente = t101.ProvinciaCliente,"
                + " s.MunicipioCliente = t101.MunicipioCliente,"
                + " s.CalleCliente = t101.CalleCliente,"
                + " s.CodPostalCliente = t101.CodPostalCliente,"
                + " s.NumeroFincaCliente = t101.NumeroFincaCliente,"
                + " s.Identificador = t101.Identificador,"
                + " s.RazonSocial = t101.RazonSocial,"
                + " s.Pais = t101.Pais,"
                + " s.Provincia = t101.Provincia,"
                + " s.Municipio = t101.Municipio,"
                + " s.Calle = t101.Calle,"
                + " s.CodPostal = t101.CodPostal,"
                + " s.NumeroFinca = t101.NumeroFinca,"
                + " s.Identificador = t101.Identificador,"
                + " s.RazonSocial = t101.RazonSocial,"
                + " s.Linea1DeLaDireccionExterna = t101.Linea1DeLaDireccionExterna,"
                + " s.Linea2DeLaDireccionExterna = t101.Linea2DeLaDireccionExterna,"
                + " s.Linea3DeLaDireccionExterna = t101.Linea3DeLaDireccionExterna,"
                + " s.Linea4DeLaDireccionExterna = t101.Linea4DeLaDireccionExterna"
                + " WHERE s.IndicadorTipoDireccion = 'F' AND"
                + " (s.Pais = '' AND s.Provincia = '' AND s.Municipio = '' AND s.Poblacion = '')"
                + " OR s.Linea1DeLaDireccionExterna = ''";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "UPDATE eexxi_solicitudes_tmp s"
                + " INNER JOIN eexxi_solicitudes_t101 t101 ON"
                + " t101.CodigoDeSolicitud = s.CodigoDeSolicitud AND"
                + " t101.CUPS = s.CUPS"
                + " SET s.PaisCliente = t101.PaisCliente," 
                + " s.ProvinciaCliente = t101.ProvinciaCliente," 
                + " s.MunicipioCliente = t101.MunicipioCliente," 
                + " s.CalleCliente = t101.CalleCliente,"
                + " s.CodPostalCliente = t101.CodPostalCliente," 
                + " s.NumeroFincaCliente = t101.NumeroFincaCliente,"
                + " s.Identificador = t101.Identificador," 
                + " s.RazonSocial = t101.RazonSocial,"
                + " s.Pais = t101.Pais," 
                + " s.Provincia = t101.Provincia," 
                + " s.Municipio = t101.Municipio,"
                + " s.Calle = t101.Calle,"
                + " s.CodPostal = t101.CodPostal,"
                + " s.NumeroFinca = t101.NumeroFinca,"
                + " s.Identificador = t101.Identificador,"
                + " s.RazonSocial = t101.RazonSocial,"
                + " s.Linea1DeLaDireccionExterna = t101.Linea1DeLaDireccionExterna,"
                + " s.Linea2DeLaDireccionExterna = t101.Linea2DeLaDireccionExterna,"
                + " s.Linea3DeLaDireccionExterna = t101.Linea3DeLaDireccionExterna,"
                + " s.Linea4DeLaDireccionExterna = t101.Linea4DeLaDireccionExterna"
                + " WHERE s.IndicadorTipoDireccion = 'F' AND"
                + " (s.Pais = '' AND s.Provincia = '' AND s.Municipio = '' AND s.Poblacion = '')"
                + " OR s.Linea1DeLaDireccionExterna = ''";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public EndesaEntity.contratacion.xxi.XML_Datos BuscaDatosSolicitudXML(string tabla, string cups22, string codigoSolicitud)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            EndesaEntity.contratacion.xxi.XML_Datos c = new EndesaEntity.contratacion.xxi.XML_Datos();

            strSql = "SELECT s.*, p.DesProvincia, p2.DesProvincia as DesProvinciaCliente "
                + " FROM " + tabla + " s"
                + " INNER JOIN eexxi_param_solicitudes_codigos c ON"
                + " c.CodigoDelProceso = s.CodigoDelProceso AND"
                + " c.CodigoDePaso = s.CodigoDePaso"
                + " LEFT OUTER JOIN eexxi_param_provincias p ON"
                + " p.CodigoPostal = s.Provincia"
                + " LEFT OUTER JOIN eexxi_param_provincias p2 ON"
                + " p2.CodigoPostal = s.ProvinciaCliente"
                + " WHERE s.CUPS = '" + cups22 + "' and"
                + " CodigoDeSolicitud = '" + codigoSolicitud + "' and"
                + " c.Descripcion = 'ALTA';";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {

                #region Campos

                c.encontrado_registro = true;

                if (r["CodigoREEEmpresaEmisora"] != System.DBNull.Value)
                    c.codigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();

                if (r["CodigoREEEmpresaDestino"] != System.DBNull.Value)
                    c.codigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();

                if (r["CodigoDelProceso"] != System.DBNull.Value)
                    c.codigoDelProceso = r["CodigoDelProceso"].ToString();

                if (r["CodigoDePaso"] != System.DBNull.Value)
                    c.codigoDePaso = r["CodigoDePaso"].ToString();

                if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                    c.codigoDeSolicitud = r["CodigoDeSolicitud"].ToString();

                if (r["SecuencialDeSolicitud"] != System.DBNull.Value)
                    c.secuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();

                if (r["FechaSolicitud"] != System.DBNull.Value)
                    c.fechaSolicitud = Convert.ToDateTime(r["FechaSolicitud"]);

                if (r["CUPS"] != System.DBNull.Value)
                    c.cups = r["CUPS"].ToString();

                if (r["CNAE"] != System.DBNull.Value)
                    c.cnae = r["CNAE"].ToString();

                if (r["PotenciaExtension"] != System.DBNull.Value)
                    c.potenciaExtension = Convert.ToInt32(r["PotenciaExtension"]);

                if (r["PotenciaDeAcceso"] != System.DBNull.Value)
                    c.potenciaDeAcceso = Convert.ToInt32(r["PotenciaDeAcceso"]);

                if (r["PotenciaInstAT"] != System.DBNull.Value)
                    c.potenciaInstAT = Convert.ToInt32(r["PotenciaInstAT"]);

                if (r["IndicativoDeInterrumpibilidad"] != System.DBNull.Value)
                    c.indicativoDeInterrumpibilidad = r["IndicativoDeInterrumpibilidad"].ToString();

                if (r["Pais"] != System.DBNull.Value)
                    c.pais = r["Pais"].ToString();

                if (r["Provincia"] != System.DBNull.Value)
                    c.provincia = r["Provincia"].ToString();

                if (r["DesProvincia"] != System.DBNull.Value)
                    c.provincia = r["DesProvincia"].ToString();

                if (r["Municipio"] != System.DBNull.Value)
                    c.municipio = r["Municipio"].ToString();

                if (r["Poblacion"] != System.DBNull.Value)
                    c.poblacion = r["Poblacion"].ToString();

                if (r["DescripcionPoblacion"] != System.DBNull.Value)
                    c.descripcionPoblacion = r["DescripcionPoblacion"].ToString();

                if (r["TipoVia"] != System.DBNull.Value)
                    c.tipoVia = r["TipoVia"].ToString();

                if (r["CodPostal"] != System.DBNull.Value)
                    c.codPostal = r["CodPostal"].ToString();

                if (r["Calle"] != System.DBNull.Value)
                    c.calle = r["Calle"].ToString();

                if (r["NumeroFinca"] != System.DBNull.Value)
                    c.numeroFinca = r["NumeroFinca"].ToString();

                if (r["AclaradorFinca"] != System.DBNull.Value)
                    c.aclaradorFinca = r["AclaradorFinca"].ToString();

                if (r["TipoIdentificador"] != System.DBNull.Value)
                    c.tipoIdentificador = r["TipoIdentificador"].ToString();

                if (r["Identificador"] != System.DBNull.Value)
                    c.identificador = r["Identificador"].ToString();

                if (r["TipoPersona"] != System.DBNull.Value)
                    c.tipoPersona = r["TipoPersona"].ToString();

                if (r["RazonSocial"] != System.DBNull.Value)
                    c.razonSocial = r["RazonSocial"].ToString();

                if (r["PrefijoPais"] != System.DBNull.Value)
                    c.prefijoPais = r["PrefijoPais"].ToString();

                if (r["Numero"] != System.DBNull.Value)
                    c.numero = r["Numero"].ToString();

                if (r["CorreoElectronico"] != System.DBNull.Value)
                    c.correoElectronico = r["CorreoElectronico"].ToString();

                if (r["IndicadorTipoDireccion"] != System.DBNull.Value)
                    c.indicadorTipoDireccion = r["IndicadorTipoDireccion"].ToString();

                if (r["IndicativoDeDireccionExterna"] != System.DBNull.Value)
                    c.indicativoDeDireccionExterna = r["IndicativoDeDireccionExterna"].ToString();

                if (r["Linea1DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea1DeLaDireccionExterna = r["Linea1DeLaDireccionExterna"].ToString();

                if (r["Linea2DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea2DeLaDireccionExterna = r["Linea2DeLaDireccionExterna"].ToString();

                if (r["Linea3DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea3DeLaDireccionExterna = r["Linea3DeLaDireccionExterna"].ToString();

                if (r["Linea4DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea4DeLaDireccionExterna = r["Linea4DeLaDireccionExterna"].ToString();

                if (r["Idioma"] != System.DBNull.Value)
                    c.idioma = r["Idioma"].ToString();

                if (r["FechaActivacion"] != System.DBNull.Value)
                    c.fechaActivacion = Convert.ToDateTime(r["FechaActivacion"]);

                if (r["CodContrato"] != System.DBNull.Value)
                    c.codContrato = r["CodContrato"].ToString();

                if (r["TipoAutoconsumo"] != System.DBNull.Value)
                    c.tipoAutoconsumo = r["TipoAutoconsumo"].ToString();

                if (r["TipoContratoATR"] != System.DBNull.Value)
                    c.tipoContratoATR = r["TipoContratoATR"].ToString();

                if (r["TarifaATR"] != System.DBNull.Value)
                    c.tarifaATR = r["TarifaATR"].ToString();

                if (r["PeriodicidadFacturacion"] != System.DBNull.Value)
                    c.periodicidadFacturacion = r["PeriodicidadFacturacion"].ToString();

                if (r["TipodeTelegestion"] != System.DBNull.Value)
                    c.tipodeTelegestion = r["TipodeTelegestion"].ToString();

                if (r["PotenciaPeriodo1"] != System.DBNull.Value)
                    c.potenciaPeriodo[1] = Convert.ToDouble(r["PotenciaPeriodo1"]);
                if (r["PotenciaPeriodo2"] != System.DBNull.Value)
                    c.potenciaPeriodo[2] = Convert.ToDouble(r["PotenciaPeriodo2"]);
                if (r["PotenciaPeriodo3"] != System.DBNull.Value)
                    c.potenciaPeriodo[3] = Convert.ToDouble(r["PotenciaPeriodo3"]);
                if (r["PotenciaPeriodo4"] != System.DBNull.Value)
                    c.potenciaPeriodo[4] = Convert.ToDouble(r["PotenciaPeriodo4"]);
                if (r["PotenciaPeriodo5"] != System.DBNull.Value)
                    c.potenciaPeriodo[5] = Convert.ToDouble(r["PotenciaPeriodo5"]);
                if (r["PotenciaPeriodo6"] != System.DBNull.Value)
                    c.potenciaPeriodo[6] = Convert.ToDouble(r["PotenciaPeriodo6"]);

                if (r["MarcaMedidaConPerdidas"] != System.DBNull.Value)
                    c.marcaMedidaConPerdidas = r["MarcaMedidaConPerdidas"].ToString();

                if (r["TensionDelSuministro"] != System.DBNull.Value)
                    c.tensionDelSuministro = Convert.ToInt32(r["TensionDelSuministro"]);

                if (r["PaisCliente"] != System.DBNull.Value)
                    c.paisCliente = r["PaisCliente"].ToString();
                if (r["ProvinciaCliente"] != System.DBNull.Value)
                    c.provinciaCliente = provincias.DesProvincia(r["ProvinciaCliente"].ToString());
                if (r["MunicipioCliente"] != System.DBNull.Value)
                    c.municipioCliente = r["MunicipioCliente"].ToString();
                if (r["PoblacionCliente"] != System.DBNull.Value)
                    c.poblacionCliente = r["PoblacionCliente"].ToString();
                if (r["DescripcionPoblacionCliente"] != System.DBNull.Value)
                    c.descripcionPoblacionCliente = r["DescripcionPoblacionCliente"].ToString();
                if (r["CodPostalCliente"] != System.DBNull.Value)
                    c.codPostalCliente = r["CodPostalCliente"].ToString();
                if (r["CalleCliente"] != System.DBNull.Value)
                    c.calleCliente = r["CalleCliente"].ToString();
                if (r["NumeroFincaCliente"] != System.DBNull.Value)
                    c.numeroFincaCliente = r["NumeroFincaCliente"].ToString();
                if (r["PisoCliente"] != System.DBNull.Value)
                    c.pisoCliente = r["PisoCliente"].ToString();


                #endregion

            }
            db.CloseConnection();
            return c;
        }

        public Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> BuscaDatosSolicitudXML(string tabla, string codigoProceso, string codigoPaso,
            Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> dic_solicitud_cups)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;



            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> d =
                new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();

            strSql = "SELECT s.*, p.DesProvincia, p2.DesProvincia as DesProvinciaCliente "
                + " FROM " + tabla + " s"
                + " LEFT OUTER JOIN eexxi_param_provincias p ON"
                + " p.CodigoPostal = s.Provincia"
                + " LEFT OUTER JOIN eexxi_param_provincias p2 ON"
                + " p2.CodigoPostal = s.ProvinciaCliente"
                + " WHERE s.CodigoDeSolicitud in";

            foreach(KeyValuePair<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> p in dic_solicitud_cups)
            {
                if (firstOnly)
                {
                    strSql += " ('" + p.Value.solicitud + "'";
                    firstOnly = false;
                }
                else
                {
                    strSql += " ,'" + p.Value.solicitud + "'";
                }
            }

            firstOnly = true;

            strSql += ") and s.CUPS in";

            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> p in dic_solicitud_cups)
            {
                if (firstOnly)
                {
                    strSql += " ('" + p.Value.cups + "'";
                    firstOnly = false;
                }
                else
                {
                    strSql += " ,'" + p.Value.cups + "'";
                }
            }


            //for (int i = 0; i < listaCodigosSolicitud.Count; i++)
            //{
            //    if (firstOnly)
            //    {
            //        strSql += " ('" + listaCodigosSolicitud[i] + "'";
            //        firstOnly = false;
            //    }
            //    else
            //    {
            //        strSql += " ,'" + listaCodigosSolicitud[i] + "'";
            //    }
            //}

            strSql += ") and CodigoDelProceso = '" + codigoProceso + "'"
                + " AND CodigoDePaso = '" + codigoPaso + "'";


            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.contratacion.xxi.XML_Datos c = new EndesaEntity.contratacion.xxi.XML_Datos();

                #region Campos



                c.encontrado_registro = true;

                if (r["CodigoREEEmpresaEmisora"] != System.DBNull.Value)
                    c.codigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();

                if (r["CodigoREEEmpresaDestino"] != System.DBNull.Value)
                    c.codigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();

                if (r["CodigoDelProceso"] != System.DBNull.Value)
                    c.codigoDelProceso = r["CodigoDelProceso"].ToString();

                if (r["CodigoDePaso"] != System.DBNull.Value)
                    c.codigoDePaso = r["CodigoDePaso"].ToString();

                if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                    c.codigoDeSolicitud = r["CodigoDeSolicitud"].ToString();

                if (r["SecuencialDeSolicitud"] != System.DBNull.Value)
                    c.secuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();

                if (r["FechaSolicitud"] != System.DBNull.Value)
                    c.fechaSolicitud = Convert.ToDateTime(r["FechaSolicitud"]);

                if (r["CUPS"] != System.DBNull.Value)
                    c.cups = r["CUPS"].ToString();

                if (r["CNAE"] != System.DBNull.Value)
                    c.cnae = r["CNAE"].ToString();

                if (r["PotenciaExtension"] != System.DBNull.Value)
                    c.potenciaExtension = Convert.ToInt32(r["PotenciaExtension"]);

                if (r["PotenciaDeAcceso"] != System.DBNull.Value)
                    c.potenciaDeAcceso = Convert.ToInt32(r["PotenciaDeAcceso"]);

                if (r["PotenciaInstAT"] != System.DBNull.Value)
                    c.potenciaInstAT = Convert.ToInt32(r["PotenciaInstAT"]);

                if (r["IndicativoDeInterrumpibilidad"] != System.DBNull.Value)
                    c.indicativoDeInterrumpibilidad = r["IndicativoDeInterrumpibilidad"].ToString();

                if (r["Pais"] != System.DBNull.Value)
                    c.pais = r["Pais"].ToString();

                if (r["Provincia"] != System.DBNull.Value)
                    c.provincia = r["Provincia"].ToString();

                if (r["DesProvincia"] != System.DBNull.Value)
                    c.provincia = r["DesProvincia"].ToString();

                if (r["Municipio"] != System.DBNull.Value)
                    c.municipio = r["Municipio"].ToString();

                if (r["Poblacion"] != System.DBNull.Value)
                    c.poblacion = r["Poblacion"].ToString();

                if (r["DescripcionPoblacion"] != System.DBNull.Value)
                    c.descripcionPoblacion = r["DescripcionPoblacion"].ToString();

                if (r["TipoVia"] != System.DBNull.Value)
                    c.tipoVia = r["TipoVia"].ToString();

                if (r["CodPostal"] != System.DBNull.Value)
                    c.codPostal = r["CodPostal"].ToString();

                if (r["Calle"] != System.DBNull.Value)
                    c.calle = r["Calle"].ToString();

                if (r["NumeroFinca"] != System.DBNull.Value)
                    c.numeroFinca = r["NumeroFinca"].ToString();

                if (r["AclaradorFinca"] != System.DBNull.Value)
                    c.aclaradorFinca = r["AclaradorFinca"].ToString();

                if (r["TipoIdentificador"] != System.DBNull.Value)
                    c.tipoIdentificador = r["TipoIdentificador"].ToString();

                if (r["Identificador"] != System.DBNull.Value)
                    c.identificador = r["Identificador"].ToString();

                if (r["TipoPersona"] != System.DBNull.Value)
                    c.tipoPersona = r["TipoPersona"].ToString();

                if (r["RazonSocial"] != System.DBNull.Value)
                    c.razonSocial = r["RazonSocial"].ToString();

                if (r["PrefijoPais"] != System.DBNull.Value)
                    c.prefijoPais = r["PrefijoPais"].ToString();

                if (r["Numero"] != System.DBNull.Value)
                    c.numero = r["Numero"].ToString();

                if (r["CorreoElectronico"] != System.DBNull.Value)
                    c.correoElectronico = r["CorreoElectronico"].ToString();

                if (r["IndicadorTipoDireccion"] != System.DBNull.Value)
                    c.indicadorTipoDireccion = r["IndicadorTipoDireccion"].ToString();

                if (r["IndicativoDeDireccionExterna"] != System.DBNull.Value)
                    c.indicativoDeDireccionExterna = r["IndicativoDeDireccionExterna"].ToString();

                if (r["Linea1DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea1DeLaDireccionExterna = r["Linea1DeLaDireccionExterna"].ToString();

                if (r["Linea2DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea2DeLaDireccionExterna = r["Linea2DeLaDireccionExterna"].ToString();

                if (r["Linea3DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea3DeLaDireccionExterna = r["Linea3DeLaDireccionExterna"].ToString();

                if (r["Linea4DeLaDireccionExterna"] != System.DBNull.Value)
                    c.linea4DeLaDireccionExterna = r["Linea4DeLaDireccionExterna"].ToString();

                if (r["Idioma"] != System.DBNull.Value)
                    c.idioma = r["Idioma"].ToString();

                if (r["FechaActivacion"] != System.DBNull.Value)
                    c.fechaActivacion = Convert.ToDateTime(r["FechaActivacion"]);

                if (r["CodContrato"] != System.DBNull.Value)
                    c.codContrato = r["CodContrato"].ToString();

                if (r["TipoAutoconsumo"] != System.DBNull.Value)
                    c.tipoAutoconsumo = r["TipoAutoconsumo"].ToString();

                if (r["TipoContratoATR"] != System.DBNull.Value)
                    c.tipoContratoATR = r["TipoContratoATR"].ToString();

                if (r["TarifaATR"] != System.DBNull.Value)
                    c.tarifaATR = r["TarifaATR"].ToString();

                if (r["PeriodicidadFacturacion"] != System.DBNull.Value)
                    c.periodicidadFacturacion = r["PeriodicidadFacturacion"].ToString();

                if (r["TipodeTelegestion"] != System.DBNull.Value)
                    c.tipodeTelegestion = r["TipodeTelegestion"].ToString();

                if (r["PotenciaPeriodo1"] != System.DBNull.Value)
                    c.potenciaPeriodo[1] = Convert.ToDouble(r["PotenciaPeriodo1"]);
                if (r["PotenciaPeriodo2"] != System.DBNull.Value)
                    c.potenciaPeriodo[2] = Convert.ToDouble(r["PotenciaPeriodo2"]);
                if (r["PotenciaPeriodo3"] != System.DBNull.Value)
                    c.potenciaPeriodo[3] = Convert.ToDouble(r["PotenciaPeriodo3"]);
                if (r["PotenciaPeriodo4"] != System.DBNull.Value)
                    c.potenciaPeriodo[4] = Convert.ToDouble(r["PotenciaPeriodo4"]);
                if (r["PotenciaPeriodo5"] != System.DBNull.Value)
                    c.potenciaPeriodo[5] = Convert.ToDouble(r["PotenciaPeriodo5"]);
                if (r["PotenciaPeriodo6"] != System.DBNull.Value)
                    c.potenciaPeriodo[6] = Convert.ToDouble(r["PotenciaPeriodo6"]);

                if (r["MarcaMedidaConPerdidas"] != System.DBNull.Value)
                    c.marcaMedidaConPerdidas = r["MarcaMedidaConPerdidas"].ToString();

                if (r["TensionDelSuministro"] != System.DBNull.Value)
                    c.tensionDelSuministro = Convert.ToInt32(r["TensionDelSuministro"]);

                if (r["PaisCliente"] != System.DBNull.Value)
                    c.paisCliente = r["PaisCliente"].ToString();
                if (r["ProvinciaCliente"] != System.DBNull.Value)
                    c.provinciaCliente = provincias.DesProvincia(r["ProvinciaCliente"].ToString());
                if (r["MunicipioCliente"] != System.DBNull.Value)
                    c.municipioCliente = r["MunicipioCliente"].ToString();
                if (r["PoblacionCliente"] != System.DBNull.Value)
                    c.poblacionCliente = r["PoblacionCliente"].ToString();
                if (r["DescripcionPoblacionCliente"] != System.DBNull.Value)
                    c.descripcionPoblacionCliente = r["DescripcionPoblacionCliente"].ToString();
                if (r["CodPostalCliente"] != System.DBNull.Value)
                    c.codPostalCliente = r["CodPostalCliente"].ToString();
                if (r["CalleCliente"] != System.DBNull.Value)
                    c.calleCliente = r["CalleCliente"].ToString();
                if (r["NumeroFincaCliente"] != System.DBNull.Value)
                    c.numeroFincaCliente = r["NumeroFincaCliente"].ToString();
                if (r["PisoCliente"] != System.DBNull.Value)
                    c.pisoCliente = r["PisoCliente"].ToString();


                #endregion
                d.Add(c.codigoDeSolicitud, c);

            }
            db.CloseConnection();
            return d;
        }

        public void ImprimeContratos(Dictionary<string, string> dic)
        {

            EndesaBusiness.pdf.JuntaHojasPDF pdfmerge = new pdf.JuntaHojasPDF();

            //FileInfo nombreArchivo = new FileInfo(param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) 
            //    + @"\" + param.GetValue("nombre_fichero_pdf_merge", DateTime.Now, DateTime.Now));

            Casos caso = new Casos();
            string portada = "";
            string carta = "";
            string contrato = "";
            string documentos_generados = "";
            string pdf_final_todos_juntos = "";
            bool firstOnly = true;
            //string parametro = "";
            EndesaEntity.contratacion.xxi.XML_Datos xml;
            List<string> lista_archivos = new List<string>();

            //if (nombreArchivo.Exists)
            //    nombreArchivo.Delete();

            pdf_final_todos_juntos = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                + @"\PDF_EEXXI_TODOS_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

            try
            {
                foreach (KeyValuePair<string, string> p in dic)
                {
                    lista_archivos.Clear();

                    xml = BuscaDatosSolicitudXML("eexxi_solicitudes_tmp", p.Key, p.Value);
                    if (!xml.encontrado_registro)
                        xml = BuscaDatosSolicitudXML("eexxi_solicitudes", p.Key, p.Value);

                    portada = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                        + "\\Caratula_" + xml.cups
                        + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

                    lista_archivos.Add(portada);

                    carta = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                        + "\\CARTA_" + xml.cups + "_"
                        + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

                    lista_archivos.Add(carta);

                    contrato = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                        + "\\CONTRATO_" + xml.cups + "_"
                        + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

                    lista_archivos.Add(contrato);


                    if (firstOnly)
                    {
                        documentos_generados += portada + "," + carta + "," + contrato;
                        firstOnly = false;
                    }
                    else
                        documentos_generados += "," + portada + "," + carta + "," + contrato;

                    CreaPortada(xml, portada);
                    CreaCarta(xml, carta);
                    CreaContrato(xml, contrato);


                    pdfmerge.MergeFiles(param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                        + @"\" + xml.cups + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now)
                        , lista_archivos);



                    //caso.MarcaComoImpreso(xml.cups, xml.codigoDeSolicitud);
                    caso.MarcaComoImpreso(xml.cups);

                    for (int i = 0; i < lista_archivos.Count(); i++)
                    {
                        FileInfo fichero = new FileInfo(lista_archivos[i]);
                        if (fichero.Exists)
                            fichero.Delete();
                    }

                }
            }
            catch(Exception ex)
            {

            }

            
            

            
            //sw.WriteLine(documentos_generados);
            //sw.Close();            

            //parametro = " -l " + nombreArchivo.FullName
            //    + " -o "                
            //    + pdf_final_todos_juntos
            //    + " concat ";

            //utilidades.Fichero.EjecutaComando(param.GetValue("pdf_merge", DateTime.Now, DateTime.Now), parametro);

            // string[] lista_archivos = documentos_generados.Split(',');
            
           


        }
        public void ImprimeContratosAgrupadosNIF(Dictionary<string, List<EndesaEntity.contratacion.xxi.Cups_Solicitud>> dic)
        {

            EndesaBusiness.pdf.JuntaHojasPDF pdfmerge = new pdf.JuntaHojasPDF();

            //FileInfo nombreArchivo = new FileInfo(param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) 
            //    + @"\" + param.GetValue("nombre_fichero_pdf_merge", DateTime.Now, DateTime.Now));

            Casos caso = new Casos();
            string portada = "";
            string carta = "";
            string contrato = "";
            string condiciones = "";
            string pdf_final_todos_juntos = "";
            bool firstOnly = true;
            //string parametro = "";
            EndesaEntity.contratacion.xxi.XML_Datos xml;
            List<string> lista_archivos = new List<string>();
            string lista_cups = "";

            //if (nombreArchivo.Exists)
            //    nombreArchivo.Delete();

            pdf_final_todos_juntos = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                + @"\PDF_EEXXI_TODOS_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

            // StreamWriter sw = new StreamWriter(nombreArchivo.FullName, true);
            foreach (KeyValuePair<string, List<EndesaEntity.contratacion.xxi.Cups_Solicitud>> p in dic)
            {
                // Para la carta ---> lista de cups
                firstOnly = true;
                for (int j = 0; j < p.Value.Count(); j++)
                {
                    xml = BuscaDatosSolicitudXML("eexxi_solicitudes_tmp", p.Value[j].cups, p.Value[j].solicitud);
                    if (!xml.encontrado_registro)
                        xml = BuscaDatosSolicitudXML("eexxi_solicitudes", p.Value[j].cups, p.Value[j].solicitud);

                    if (firstOnly)
                    {
                        lista_cups = xml.cups;
                        firstOnly = false;
                    }
                    else
                    {
                        lista_cups = lista_cups + ", " + xml.cups;
                    }
                    
                }


                firstOnly = true;
                for (int j = 0; j < p.Value.Count() -1; j++)
                {

                    xml = BuscaDatosSolicitudXML("eexxi_solicitudes_tmp", p.Value[j].cups, p.Value[j].solicitud);
                    if (!xml.encontrado_registro)
                        xml = BuscaDatosSolicitudXML("eexxi_solicitudes", p.Value[j].cups, p.Value[j].solicitud);
                    

                    if (firstOnly)
                    {                       

                        portada = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                            + "\\Caratula_" + p.Key
                            + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

                        lista_archivos.Add(portada);

                        carta = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                            + "\\CARTA_" + p.Key + "_"
                            + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

                        lista_archivos.Add(carta);
                        CreaPortadaAgrupada(xml, portada);
                        CreaCartaAgrupada(xml, carta, lista_cups);
                        firstOnly = false;
                    }

                    
                    contrato = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                    + "\\CONTRATO_" + xml.cups + "_"
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

                    lista_archivos.Add(contrato);                                      
                    CreaContratoAgrupada(xml, contrato);
                   
                }
                

                int w = p.Value.Count() -1;
                xml = BuscaDatosSolicitudXML("eexxi_solicitudes_tmp", p.Value[w].cups, p.Value[w].solicitud);
                if (!xml.encontrado_registro)
                    xml = BuscaDatosSolicitudXML("eexxi_solicitudes", p.Value[w].cups, p.Value[w].solicitud);

                contrato = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                    + "\\CONTRATO_" + xml.cups + "_"
                    + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);

                lista_archivos.Add(contrato);
                CreaContrato(xml, contrato);

                //condiciones = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                //            + "\\Condiciones_" + p.Key
                //            + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now);
                //lista_archivos.Add(condiciones);
                //CreaCondiciones(condiciones);
                // Para la carta ---> lista de cups

                for (int j = 0; j < p.Value.Count(); j++)
                {
                    xml = BuscaDatosSolicitudXML("eexxi_solicitudes_tmp", p.Value[j].cups, p.Value[j].solicitud);
                    if (!xml.encontrado_registro)
                        xml = BuscaDatosSolicitudXML("eexxi_solicitudes", p.Value[j].cups, p.Value[j].solicitud);

                    //caso.MarcaComoImpreso(xml.cups, xml.codigoDeSolicitud);
                    caso.MarcaComoImpreso(xml.cups);

                }

                pdfmerge.MergeFiles(param.GetValue("salida_contrato", DateTime.Now, DateTime.Now)
                    + @"\" + p.Key + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") 
                    + param.GetValue("extension_pdf", DateTime.Now, DateTime.Now)
                    ,lista_archivos);
                                

                for (int i = 0; i < lista_archivos.Count(); i++)
                {
                    FileInfo fichero = new FileInfo(lista_archivos[i]);
                    if (fichero.Exists)
                        fichero.Delete();
                }


                for (int j = 0; j < p.Value.Count(); j++)                   
                    caso.MarcaComoImpreso(xml.cups);

                

            }


            //sw.WriteLine(documentos_generados);
            //sw.Close();            

            //parametro = " -l " + nombreArchivo.FullName
            //    + " -o "                
            //    + pdf_final_todos_juntos
            //    + " concat ";

            //utilidades.Fichero.EjecutaComando(param.GetValue("pdf_merge", DateTime.Now, DateTime.Now), parametro);

            // string[] lista_archivos = documentos_generados.Split(',');




        }
        private void CreaPortada(EndesaEntity.contratacion.xxi.XML_Datos xml, string ruta_fichero)
        {
            string template_normal = System.Environment.CurrentDirectory + param.GetValue("plantilla_caratula", DateTime.Now, DateTime.Now);
            FileInfo fichero = new FileInfo(template_normal);

            if (fichero.Exists)
            {
                // string salida = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) + "\\Caratula_" + xml.cups + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf";


                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
                Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(template_normal, ReadOnly: false, Visible: false);


                if (aDoc.Bookmarks.Exists("razon_social"))
                {
                    aDoc.Bookmarks["razon_social"].Select();
                    wordApp.Selection.TypeText(xml.razonSocial);
                }

                if (aDoc.Bookmarks.Exists("dir1"))
                {
                    aDoc.Bookmarks["dir1"].Select();
                    if (xml.linea1DeLaDireccionExterna != null)
                        wordApp.Selection.TypeText(xml.linea1DeLaDireccionExterna + " " + xml.linea2DeLaDireccionExterna);
                    else
                        wordApp.Selection.TypeText(xml.tipoViaCliente + " " + xml.calleCliente + " " + xml.numeroFincaCliente);
                }

                if (aDoc.Bookmarks.Exists("dir_2"))
                {
                    aDoc.Bookmarks["dir_2"].Select();
                    if (xml.linea1DeLaDireccionExterna != null)
                        wordApp.Selection.TypeText(xml.linea3DeLaDireccionExterna + " " + xml.linea4DeLaDireccionExterna);
                    else
                        wordApp.Selection.TypeText(xml.descripcionPoblacionCliente + ", " + xml.codPostalCliente + " " + xml.provinciaCliente);
                }


                wordApp.ActiveDocument.ExportAsFixedFormat(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
                Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);
                aDoc.Close(false);
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
                Marshal.ReleaseComObject(aDoc);
                wordApp = null;
                aDoc = null;
            }
            else
            {
                MessageBox.Show("No existe el archivo "
                    + fichero.FullName + "."
                    + System.Environment.NewLine
                    + "No se puede imprmir el contrato",
               "EEXXI - CreaPortada",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);

            }

        }
        private void CreaPortadaAgrupada(EndesaEntity.contratacion.xxi.XML_Datos xml, string ruta_fichero)
        {
            string template_normal = System.Environment.CurrentDirectory + param.GetValue("plantilla_caratula", DateTime.Now, DateTime.Now);
            FileInfo fichero = new FileInfo(template_normal);

            if (fichero.Exists)
            {
                // string salida = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) + "\\Caratula_" + xml.cups + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf";


                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
                Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(template_normal, ReadOnly: false, Visible: false);


                if (aDoc.Bookmarks.Exists("razon_social"))
                {
                    aDoc.Bookmarks["razon_social"].Select();
                    wordApp.Selection.TypeText(xml.razonSocial);
                }

                if (aDoc.Bookmarks.Exists("dir1"))
                {
                    aDoc.Bookmarks["dir1"].Select();
                    if (xml.linea1DeLaDireccionExterna != null)
                        wordApp.Selection.TypeText(xml.linea1DeLaDireccionExterna + " " + xml.linea2DeLaDireccionExterna);
                    else
                        wordApp.Selection.TypeText(xml.tipoViaCliente + " " + xml.calleCliente + " " + xml.numeroFincaCliente);
                }

                if (aDoc.Bookmarks.Exists("dir_2"))
                {
                    aDoc.Bookmarks["dir_2"].Select();
                    if (xml.linea1DeLaDireccionExterna != null)
                        wordApp.Selection.TypeText(xml.linea3DeLaDireccionExterna + " " + xml.linea4DeLaDireccionExterna);
                    else
                        wordApp.Selection.TypeText(xml.descripcionPoblacionCliente + ", " + xml.codPostalCliente + " " + xml.provinciaCliente);
                }

                //if (aDoc.Bookmarks.Exists("dir1"))
                //{
                //    aDoc.Bookmarks["dir1"].Select();
                //    wordApp.Selection.TypeText(xml.tipoViaCliente + " " + xml.calleCliente + " " + xml.numeroFincaCliente);
                //}

                //if (aDoc.Bookmarks.Exists("dir_2"))
                //{
                //    aDoc.Bookmarks["dir_2"].Select();
                //    if(xml.descripcionPoblacionCliente == "" && xml.provinciaCliente == "")
                //    {
                //        wordApp.Selection.TypeText(xml.linea2DeLaDireccionExterna + ", " + xml.linea3DeLaDireccionExterna);
                //    }
                //    else
                //        wordApp.Selection.TypeText(xml.descripcionPoblacionCliente + ", " + xml.codPostalCliente + " " + xml.provinciaCliente); 
                //}


                wordApp.ActiveDocument.ExportAsFixedFormat(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
                Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);
                aDoc.Close(false);
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
                Marshal.ReleaseComObject(aDoc);
                wordApp = null;
                aDoc = null;
            }
            else
            {
                MessageBox.Show("No existe el archivo "
                    + fichero.FullName + "."
                    + System.Environment.NewLine
                    + "No se puede imprmir el contrato",
               "EEXXI - CreaPortada",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);

            }

        }
        private void CreaCarta(EndesaEntity.contratacion.xxi.XML_Datos xml, string ruta_fichero)
        {

            string template_normal = System.Environment.CurrentDirectory + param.GetValue("plantilla_carta", DateTime.Now, DateTime.Now);
            //string salida = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) + "\\CARTA_" + xml.cups + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf";


            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(template_normal, ReadOnly: false, Visible: false);

            if (aDoc.Bookmarks.Exists("FD"))
            {
                aDoc.Bookmarks["FD"].Select();
                wordApp.Selection.TypeText(xml.fechaActivacion.ToString("dd/MM/yyyy"));
            }


            if (aDoc.Bookmarks.Exists("ApelllidosNombre"))
            {
                aDoc.Bookmarks["ApelllidosNombre"].Select();
                wordApp.Selection.TypeText(xml.razonSocial);
            }

            if (aDoc.Bookmarks.Exists("Direccion"))
            {
                aDoc.Bookmarks["Direccion"].Select();
                if (xml.linea1DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea1DeLaDireccionExterna + " " + xml.linea2DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.tipoViaCliente + " " + xml.calleCliente + " " + xml.numeroFincaCliente);
            }

            if (aDoc.Bookmarks.Exists("CPLocalidad"))
            {
                aDoc.Bookmarks["CPLocalidad"].Select();
                if (xml.linea1DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea3DeLaDireccionExterna + " " + xml.linea4DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.descripcionPoblacionCliente + ", " + xml.codPostalCliente + " " + xml.provinciaCliente);
            }

            if (aDoc.Bookmarks.Exists("CUPS_LARGO"))
            {
                aDoc.Bookmarks["CUPS_LARGO"].Select();
                wordApp.Selection.TypeText(xml.cups);
            }


            wordApp.ActiveDocument.ExportAsFixedFormat(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
                Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);
            aDoc.Close(false);            
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            Marshal.ReleaseComObject(aDoc);          
            wordApp = null;
            aDoc = null;

        }
        private void CreaCartaAgrupada(EndesaEntity.contratacion.xxi.XML_Datos xml, string ruta_fichero, string lista_cups)
        {

            string template_normal = System.Environment.CurrentDirectory + param.GetValue("plantilla_carta_agrupada", DateTime.Now, DateTime.Now);
            //string salida = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) + "\\CARTA_" + xml.cups + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf";


            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(template_normal, ReadOnly: false, Visible: false);

            if (aDoc.Bookmarks.Exists("FD"))
            {
                aDoc.Bookmarks["FD"].Select();
                wordApp.Selection.TypeText(xml.fechaActivacion.ToString("dd/MM/yyyy"));
            }


            if (aDoc.Bookmarks.Exists("ApelllidosNombre"))
            {
                aDoc.Bookmarks["ApelllidosNombre"].Select();
                if (xml.razonSocial != null)
                    wordApp.Selection.TypeText(xml.razonSocial);
            }

            if (aDoc.Bookmarks.Exists("Direccion"))
            {
                aDoc.Bookmarks["Direccion"].Select();
                if (xml.linea1DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea1DeLaDireccionExterna + " " + xml.linea2DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.tipoViaCliente + " " + xml.calleCliente + " " + xml.numeroFincaCliente);
            }

            if (aDoc.Bookmarks.Exists("CPLocalidad"))
            {
                aDoc.Bookmarks["CPLocalidad"].Select();
                if (xml.linea1DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea3DeLaDireccionExterna + " " + xml.linea4DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.descripcionPoblacionCliente + ", " + xml.codPostalCliente + " " + xml.provinciaCliente);
            }

            if (aDoc.Bookmarks.Exists("CUPS_LARGO"))
            {
                aDoc.Bookmarks["CUPS_LARGO"].Select();
                if (lista_cups != null)
                    wordApp.Selection.TypeText(lista_cups);
            }


            wordApp.ActiveDocument.ExportAsFixedFormat(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
                Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);
            aDoc.Close(false);
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            Marshal.ReleaseComObject(aDoc);
            wordApp = null;
            aDoc = null;

        }
        private void CreaContrato(EndesaEntity.contratacion.xxi.XML_Datos xml, string ruta_fichero)
        {
            cups.TarifaATR tarifa_atr = new cups.TarifaATR();
            cups.TensionATR tension_atr = new cups.TensionATR();
            contratacion.CNAE cnae = new CNAE();

            string template_normal = System.Environment.CurrentDirectory + param.GetValue("plantilla_contrato", DateTime.Now, DateTime.Now);
            //string salida = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) + "\\CONTRATO_" + xml.cups + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf";

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(template_normal, ReadOnly: false, Visible: false);

            #region DatosDelCliente

            if (aDoc.Bookmarks.Exists("ApelllidosNombre"))
            {
                aDoc.Bookmarks["ApelllidosNombre"].Select();
                wordApp.Selection.TypeText(xml.razonSocial);
            }

            if (aDoc.Bookmarks.Exists("NIFCIF"))
            {
                aDoc.Bookmarks["NIFCIF"].Select();
                wordApp.Selection.TypeText(xml.identificador);
            }

            if (aDoc.Bookmarks.Exists("CNAE"))
            {
                aDoc.Bookmarks["CNAE"].Select();
                if (xml.cnae != null)
                    wordApp.Selection.TypeText(xml.cnae.ToString());
            }

            if (aDoc.Bookmarks.Exists("Direccion"))
            {
                aDoc.Bookmarks["Direccion"].Select();
                if (xml.linea1DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea1DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.tipoViaCliente + " " + xml.calleCliente);
            }

            if (aDoc.Bookmarks.Exists("EscPuerPiso"))
            {
                aDoc.Bookmarks["EscPuerPiso"].Select();
                if (xml.numeroFincaCliente != null)
                    wordApp.Selection.TypeText(xml.pisoCliente);
            }


            if (aDoc.Bookmarks.Exists("CPLocalidad"))
            {
                aDoc.Bookmarks["CPLocalidad"].Select();
                if (xml.linea2DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea2DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.codPostalCliente + ", " + xml.descripcionPoblacionCliente);
            }

            if (aDoc.Bookmarks.Exists("Prov"))
            {
                aDoc.Bookmarks["Prov"].Select();
                wordApp.Selection.TypeText(xml.provinciaCliente);
            }

            if (aDoc.Bookmarks.Exists("Telf1"))
            {
                aDoc.Bookmarks["Telf1"].Select();
                wordApp.Selection.TypeText(xml.numero);
            }

            if (aDoc.Bookmarks.Exists("Email"))
            {
                aDoc.Bookmarks["Email"].Select();
                wordApp.Selection.TypeText(xml.correoElectronico);
            }

            if (aDoc.Bookmarks.Exists("IDIOMA"))
            {
                aDoc.Bookmarks["IDIOMA"].Select();
                wordApp.Selection.TypeText(xml.idioma);
            }

            if (aDoc.Bookmarks.Exists("Via"))
            {
                aDoc.Bookmarks["Via"].Select();
                wordApp.Selection.TypeText(xml.calle);
            }


            if (aDoc.Bookmarks.Exists("Num"))
            {
                aDoc.Bookmarks["Num"].Select();
                wordApp.Selection.TypeText(xml.numeroFinca);
            }

            if (aDoc.Bookmarks.Exists("Esc"))
            {
                aDoc.Bookmarks["Esc"].Select();
                //wordApp.Selection.TypeText();
            }

            //if (aDoc.Bookmarks.Exists("Piso"))
            //{
            //    aDoc.Bookmarks["Piso"].Select();
            //    if (xml.aclaradorFinca != null)
            //        wordApp.Selection.TypeText(xml.aclaradorFinca);
            //}

            //if (aDoc.Bookmarks.Exists("Letra"))
            //{
            //    aDoc.Bookmarks["Letra"].Select();
            //    if (xml.aclaradorFinca != null)
            //        wordApp.Selection.TypeText(xml.aclaradorFinca);
            //}

            if (aDoc.Bookmarks.Exists("CP"))
            {
                aDoc.Bookmarks["CP"].Select();
                if (xml.codPostal != null)
                    wordApp.Selection.TypeText(xml.codPostal);
            }

            if (aDoc.Bookmarks.Exists("Localidad"))
            {
                aDoc.Bookmarks["Localidad"].Select();
                if (xml.descripcionPoblacion != null)
                    wordApp.Selection.TypeText(xml.descripcionPoblacion);
            }

            if (aDoc.Bookmarks.Exists("Provincia"))
            {
                aDoc.Bookmarks["Provincia"].Select();
                if (xml.provincia != null)
                    wordApp.Selection.TypeText(xml.provincia);
            }

            if (aDoc.Bookmarks.Exists("CUPS_LARGO"))
            {
                aDoc.Bookmarks["CUPS_LARGO"].Select();
                if (xml.cups != null)
                    wordApp.Selection.TypeText(xml.cups);
            }

            #endregion

            if (aDoc.Bookmarks.Exists("CD_TARIFA"))
            {
                aDoc.Bookmarks["CD_TARIFA"].Select();
                if (xml.tarifaATR != null)
                    wordApp.Selection.TypeText(tarifa_atr.GetDescription(xml.tarifaATR));
            }

            #region TARIFA DE ACCESO

            if (aDoc.Bookmarks.Exists("P1"))
            {
                aDoc.Bookmarks["P1"].Select();
                if (xml.potenciaPeriodo[1] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[1] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P2"))
            {
                aDoc.Bookmarks["P2"].Select();
                if (xml.potenciaPeriodo[2] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[2] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P3"))
            {
                aDoc.Bookmarks["P3"].Select();
                if (xml.potenciaPeriodo[3] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[3] / 1000));
            }

            if (aDoc.Bookmarks.Exists("P4"))
            {
                aDoc.Bookmarks["P4"].Select();
                if (xml.potenciaPeriodo[4] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[4] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P5"))
            {
                aDoc.Bookmarks["P5"].Select();
                if (xml.potenciaPeriodo[5] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[5] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P6"))
            {
                aDoc.Bookmarks["P6"].Select();
                if (xml.potenciaPeriodo[6] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[6] / 1000));

            }

            if (aDoc.Bookmarks.Exists("NM_TENSION"))
            {
                aDoc.Bookmarks["NM_TENSION"].Select();
                if (xml.tensionDelSuministro != 0)
                    wordApp.Selection.TypeText(Convert.ToString(tension_atr.GetDescription(xml.tensionDelSuministro)));
            }

            #endregion

            if (aDoc.Bookmarks.Exists("TUR"))
            {
                aDoc.Bookmarks["TUR"].Select();
                wordApp.Selection.TypeText(Convert.ToString(param.GetValue("producto_contratato", DateTime.Now, DateTime.Now)));
            }

            //aDoc.SaveAs2(salida, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            wordApp.ActiveDocument.ExportAsFixedFormat(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
                Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);
            aDoc.Close(false);
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            Marshal.ReleaseComObject(aDoc);
            wordApp = null;
            aDoc = null;


        }
        private void CreaContratoAgrupada(EndesaEntity.contratacion.xxi.XML_Datos xml, string ruta_fichero)
        {
            cups.TarifaATR tarifa_atr = new cups.TarifaATR();
            cups.TensionATR tension_atr = new cups.TensionATR();
            contratacion.CNAE cnae = new CNAE();

            string template_normal = System.Environment.CurrentDirectory + param.GetValue("plantilla_contrato_agrupado", DateTime.Now, DateTime.Now);
            //string salida = param.GetValue("salida_contrato", DateTime.Now, DateTime.Now) + "\\CONTRATO_" + xml.cups + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf";

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(template_normal, ReadOnly: false, Visible: false);

            #region DatosDelCliente

            if (aDoc.Bookmarks.Exists("ApelllidosNombre"))
            {
                aDoc.Bookmarks["ApelllidosNombre"].Select();
                wordApp.Selection.TypeText(xml.razonSocial);
            }

            if (aDoc.Bookmarks.Exists("NIFCIF"))
            {
                aDoc.Bookmarks["NIFCIF"].Select();
                wordApp.Selection.TypeText(xml.identificador);
            }

            if (aDoc.Bookmarks.Exists("CNAE"))
            {
                aDoc.Bookmarks["CNAE"].Select();
                if(xml.cnae != null)
                    wordApp.Selection.TypeText(xml.cnae.ToString());
            }

            if (aDoc.Bookmarks.Exists("Direccion"))
            {
                aDoc.Bookmarks["Direccion"].Select();
                if (xml.linea1DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea1DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.tipoViaCliente + " " + xml.calleCliente);
            }

            if (aDoc.Bookmarks.Exists("EscPuerPiso"))
            {
                aDoc.Bookmarks["EscPuerPiso"].Select();
                if (xml.numeroFincaCliente != null)
                    wordApp.Selection.TypeText(xml.pisoCliente);
            }


            if (aDoc.Bookmarks.Exists("CPLocalidad"))
            {
                aDoc.Bookmarks["CPLocalidad"].Select();
                if (xml.linea2DeLaDireccionExterna != null)
                    wordApp.Selection.TypeText(xml.linea2DeLaDireccionExterna);
                else
                    wordApp.Selection.TypeText(xml.codPostalCliente + ", " + xml.descripcionPoblacionCliente);
            }

            if (aDoc.Bookmarks.Exists("Prov"))
            {
                aDoc.Bookmarks["Prov"].Select();
                wordApp.Selection.TypeText(xml.provinciaCliente);
            }

            if (aDoc.Bookmarks.Exists("Telf1"))
            {
                aDoc.Bookmarks["Telf1"].Select();
                wordApp.Selection.TypeText(xml.numero);
            }

            if (aDoc.Bookmarks.Exists("Email"))
            {
                aDoc.Bookmarks["Email"].Select();
                wordApp.Selection.TypeText(xml.correoElectronico);
            }

            if (aDoc.Bookmarks.Exists("IDIOMA"))
            {
                aDoc.Bookmarks["IDIOMA"].Select();
                wordApp.Selection.TypeText(xml.idioma);
            }

            if (aDoc.Bookmarks.Exists("Via"))
            {
                aDoc.Bookmarks["Via"].Select();
                wordApp.Selection.TypeText(xml.calle);
            }


            if (aDoc.Bookmarks.Exists("Num"))
            {
                aDoc.Bookmarks["Num"].Select();
                wordApp.Selection.TypeText(xml.numeroFinca);
            }

            if (aDoc.Bookmarks.Exists("Esc"))
            {
                aDoc.Bookmarks["Esc"].Select();
                //wordApp.Selection.TypeText();
            }

            //if (aDoc.Bookmarks.Exists("Piso"))
            //{
            //    aDoc.Bookmarks["Piso"].Select();
            //    if (xml.aclaradorFinca != null)
            //        wordApp.Selection.TypeText(xml.aclaradorFinca);
            //}

            //if (aDoc.Bookmarks.Exists("Letra"))
            //{
            //    aDoc.Bookmarks["Letra"].Select();
            //    if (xml.aclaradorFinca != null)
            //        wordApp.Selection.TypeText(xml.aclaradorFinca);
            //}

            if (aDoc.Bookmarks.Exists("CP"))
            {
                aDoc.Bookmarks["CP"].Select();
                if (xml.codPostal != null)
                    wordApp.Selection.TypeText(xml.codPostal);
            }

            if (aDoc.Bookmarks.Exists("Localidad"))
            {
                aDoc.Bookmarks["Localidad"].Select();
                if (xml.descripcionPoblacion != null)
                    wordApp.Selection.TypeText(xml.descripcionPoblacion);
            }

            if (aDoc.Bookmarks.Exists("Provincia"))
            {
                aDoc.Bookmarks["Provincia"].Select();
                if (xml.provincia != null)
                    wordApp.Selection.TypeText(xml.provincia);
            }

            if (aDoc.Bookmarks.Exists("CUPS_LARGO"))
            {
                aDoc.Bookmarks["CUPS_LARGO"].Select();
                if (xml.cups != null)
                    wordApp.Selection.TypeText(xml.cups);
            }

            #endregion

            if (aDoc.Bookmarks.Exists("CD_TARIFA"))
            {
                aDoc.Bookmarks["CD_TARIFA"].Select();
                if (xml.tarifaATR != null)
                    wordApp.Selection.TypeText(tarifa_atr.GetDescription(xml.tarifaATR));
            }

            #region TARIFA DE ACCESO

            if (aDoc.Bookmarks.Exists("P1"))
            {
                aDoc.Bookmarks["P1"].Select();
                if (xml.potenciaPeriodo[1] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[1] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P2"))
            {
                aDoc.Bookmarks["P2"].Select();
                if (xml.potenciaPeriodo[2] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[2] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P3"))
            {
                aDoc.Bookmarks["P3"].Select();
                if (xml.potenciaPeriodo[3] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[3] / 1000));
            }

            if (aDoc.Bookmarks.Exists("P4"))
            {
                aDoc.Bookmarks["P4"].Select();
                if (xml.potenciaPeriodo[4] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[4] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P5"))
            {
                aDoc.Bookmarks["P5"].Select();
                if (xml.potenciaPeriodo[5] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[5] / 1000));

            }

            if (aDoc.Bookmarks.Exists("P6"))
            {
                aDoc.Bookmarks["P6"].Select();
                if (xml.potenciaPeriodo[6] != 0)
                    wordApp.Selection.TypeText(Convert.ToString(xml.potenciaPeriodo[6] / 1000));

            }

            if (aDoc.Bookmarks.Exists("NM_TENSION"))
            {
                aDoc.Bookmarks["NM_TENSION"].Select();
                if (xml.tensionDelSuministro != 0)
                    wordApp.Selection.TypeText(Convert.ToString(tension_atr.GetDescription(xml.tensionDelSuministro)));
            }

            #endregion

            if (aDoc.Bookmarks.Exists("TUR"))
            {
                aDoc.Bookmarks["TUR"].Select();
                wordApp.Selection.TypeText(Convert.ToString(param.GetValue("producto_contratato", DateTime.Now, DateTime.Now)));
            }

            //aDoc.SaveAs2(salida, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            wordApp.ActiveDocument.ExportAsFixedFormat(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
                Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);
            aDoc.Close(false);
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            Marshal.ReleaseComObject(aDoc);
            wordApp = null;
            aDoc = null;


        }
        private void CreaCondiciones( string ruta_fichero)
        {

            string template_normal = System.Environment.CurrentDirectory + param.GetValue("plantilla_condiciones_agrupado", DateTime.Now, DateTime.Now);            

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(template_normal, ReadOnly: false, Visible: false);
            wordApp.ActiveDocument.SaveAs2(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);             
            //wordApp.ActiveDocument.ExportAsFixedFormat(ruta_fichero, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
            //    Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);
            aDoc.Close(false);
            wordApp.Quit();
            Marshal.ReleaseComObject(wordApp);
            Marshal.ReleaseComObject(aDoc);
            wordApp = null;
            aDoc = null;

        }
        public void ProcesaSolicitudes()
        {
            try
            {
                sc = new SolicitudesCodigos();
                inventario = new Inventario();
                inventarioDetalleEstados = new InventarioDetalleEstados();

                AnalizaSolicitudes(Solicitudes("ALTA"));
                AnalizaSolicitudes(Solicitudes("BAJA"));
                CreaCasos();
                // GeneraExcelComprobacion();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "EEXXI - ProcesaSolicitudes",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }
        private void AnalizaSolicitudes(Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic)
        {

            try
            {

                foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.XML_Datos> p in dic)
                {
                    inventario.AnalizaInventario(p.Value.cups, sc.TipoProceso(p.Value.codigoDelProceso, p.Value.codigoDePaso));
                    inventarioDetalleEstados.AnalizaSolicitud(p.Value.cups, p.Value);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "EEXXI - AnalizaSolicitudes",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }

        }
                

        public void SolicitudesTMP_a_Solicitudes()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                strSql = "replace into eexxi_solicitudes select * from eexxi_solicitudes_tmp";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "EEXXI - SolicitudesTMP_a_Solicitudes",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public void Borra_tabla(string nombre_tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                strSql = "delete from " + nombre_tabla;
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "EEXXI - BorraTabla",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void Vuelca_eexxi_solicitudes_tmp_a_eexxi_solicitudes()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                strSql = "replace into eexxi_solicitudes select * from eexxi_solicitudes_tmp;";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "EEXXI - Vuelca_eexxi_solicitudes_tmp_a_eexxi_solicitudes",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public string ExtraeCUPS22_DesdeNombreFichero(string cups22)
        {
            string cups = "";
            int posicion_desde = 0;

            for (int i = 1; i <= cups22.Length; i++)
            {


                if (!cups22.Substring(posicion_desde, i).Contains("_"))
                    cups = cups22.Substring(posicion_desde, i);
                else
                {
                    posicion_desde = i > 0 ? i - 1 : 0;
                    cups = "";
                }


                if (cups.Length == 20)
                    return cups;
            }

            return cups;
        }


        public void PuntosActivos()
        {
            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_altas;
            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_bajas;

            contratacion.PS_AT_Funciones psat = new PS_AT_Funciones();
            psat.Carga_PS_AT();

            dic_altas = Solicitudes("ALTA");
            dic_bajas = Solicitudes("BAJA");

            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.XML_Datos> p in dic_altas)
            {
                EndesaEntity.contratacion.xxi.XML_Datos o;
                if (!dic_bajas.TryGetValue(p.Key, out o))
                {
                    EndesaEntity.contratacion.PS_AT c;
                    p.Value.existe_en_psat = psat.l_ps_at.TryGetValue(p.Key.Substring(0, 20), out c);
                    dic_XML.Add(p.Key, p.Value);
                }

                else
                {
                    if (o.fechaActivacion < p.Value.fechaActivacion)
                    {
                        dic_XML.Add(p.Key, p.Value);
                    }

                }
            }

            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.XML_Datos> p in dic_altas)
            {
                EndesaEntity.contratacion.xxi.XML_Datos o;
                if (!dic_bajas.TryGetValue(p.Key, out o))
                {
                    EndesaEntity.contratacion.PS_AT c;
                    p.Value.existe_en_psat = psat.l_ps_at.TryGetValue(p.Key.Substring(0, 20), out c);
                    dic_XML.Add(p.Key, p.Value);
                }

                else
                {
                    if (o.fechaActivacion < p.Value.fechaActivacion)
                    {
                        dic_XML.Add(p.Key, p.Value);
                    }

                }
            }



            }

        private Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> Solicitudes(string tipo_solicitud)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();
            try
            {
                strSql = "select CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, sol.CodigoDelProceso,"
                    + " sol.CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud, CUPS, CNAE, PotenciaExtension,"
                    + " PotenciaDeAcceso, PotenciaInstAT, IndicativoDeInterrumpibilidad, Pais, Provincia, Municipio, Poblacion,"
                    + " DescripcionPoblacion, TipoVia, CodPostal, Calle, NumeroFinca, AclaradorFinca, TipoIdentificador, Identificador,"
                    + " TipoPersona, RazonSocial, PrefijoPais, Numero, CorreoElectronico, IndicadorTipoDireccion, IndicativoDeDireccionExterna,"
                    + " Linea1DeLaDireccionExterna, Linea2DeLaDireccionExterna, Linea3DeLaDireccionExterna, Linea4DeLaDireccionExterna,"
                    + " Idioma, FechaActivacion, CodContrato, TipoAutoconsumo, TipoContratoATR, TarifaATR, PeriodicidadFacturacion, TipodeTelegestion,"
                    + " PotenciaPeriodo1, PotenciaPeriodo2, PotenciaPeriodo3, PotenciaPeriodo4, PotenciaPeriodo5, PotenciaPeriodo6, "
                    + " MarcaMedidaConPerdidas, TensionDelSuministro,"
                    + " sol.created_by, fichero"
                    + " from eexxi_solicitudes_tmp sol inner join"
                    + " eexxi_param_solicitudes_codigos cod on"
                    + " cod.CodigoDelProceso = sol.CodigoDelProceso and"
                    + " cod.CodigoDePaso = sol.CodigoDePaso";

                if (tipo_solicitud != null)
                    strSql += " where cod.Descripcion = '" + tipo_solicitud + "'";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.XML_Datos c = new EndesaEntity.contratacion.xxi.XML_Datos();

                    #region Campos

                    if (r["CodigoREEEmpresaEmisora"] != System.DBNull.Value)
                        c.codigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();

                    if (r["CodigoREEEmpresaDestino"] != System.DBNull.Value)
                        c.codigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();

                    if (r["CodigoDelProceso"] != System.DBNull.Value)
                        c.codigoDelProceso = r["CodigoDelProceso"].ToString();

                    if (r["CodigoDePaso"] != System.DBNull.Value)
                        c.codigoDePaso = r["CodigoDePaso"].ToString();

                    if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                        c.codigoDeSolicitud = r["CodigoDeSolicitud"].ToString();

                    if (r["SecuencialDeSolicitud"] != System.DBNull.Value)
                        c.secuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();

                    if (r["FechaSolicitud"] != System.DBNull.Value)
                        c.fechaSolicitud = Convert.ToDateTime(r["FechaSolicitud"]);

                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups = r["CUPS"].ToString();

                    if (r["CNAE"] != System.DBNull.Value)
                        c.cnae = r["CNAE"].ToString(); ;

                    if (r["PotenciaExtension"] != System.DBNull.Value)
                        c.potenciaExtension = Convert.ToInt32(r["PotenciaExtension"]);

                    if (r["PotenciaDeAcceso"] != System.DBNull.Value)
                        c.potenciaDeAcceso = Convert.ToInt32(r["PotenciaDeAcceso"]);

                    if (r["PotenciaInstAT"] != System.DBNull.Value)
                        c.potenciaInstAT = Convert.ToInt32(r["PotenciaInstAT"]);

                    if (r["IndicativoDeInterrumpibilidad"] != System.DBNull.Value)
                        c.indicativoDeInterrumpibilidad = r["IndicativoDeInterrumpibilidad"].ToString();

                    if (r["Pais"] != System.DBNull.Value)
                        c.pais = r["Pais"].ToString();

                    if (r["Provincia"] != System.DBNull.Value)
                        c.provincia = r["Provincia"].ToString();

                    if (r["Municipio"] != System.DBNull.Value)
                        c.municipio = r["Municipio"].ToString();

                    if (r["Poblacion"] != System.DBNull.Value)
                        c.poblacion = r["Poblacion"].ToString();

                    if (r["DescripcionPoblacion"] != System.DBNull.Value)
                        c.descripcionPoblacion = r["DescripcionPoblacion"].ToString();

                    if (r["TipoVia"] != System.DBNull.Value)
                        c.tipoVia = r["TipoVia"].ToString();

                    if (r["CodPostal"] != System.DBNull.Value)
                        c.codPostal = r["CodPostal"].ToString();

                    if (r["Calle"] != System.DBNull.Value)
                        c.calle = r["Calle"].ToString();

                    if (r["NumeroFinca"] != System.DBNull.Value)
                        c.numeroFinca = r["NumeroFinca"].ToString();

                    if (r["AclaradorFinca"] != System.DBNull.Value)
                        c.aclaradorFinca = r["AclaradorFinca"].ToString();

                    if (r["TipoIdentificador"] != System.DBNull.Value)
                        c.tipoIdentificador = r["TipoIdentificador"].ToString();

                    if (r["Identificador"] != System.DBNull.Value)
                        c.identificador = r["Identificador"].ToString();

                    if (r["TipoPersona"] != System.DBNull.Value)
                        c.tipoPersona = r["TipoPersona"].ToString();

                    if (r["RazonSocial"] != System.DBNull.Value)
                        c.razonSocial = r["RazonSocial"].ToString();

                    if (r["PrefijoPais"] != System.DBNull.Value)
                        c.prefijoPais = r["PrefijoPais"].ToString();

                    if (r["Numero"] != System.DBNull.Value)
                        c.numero = r["Numero"].ToString();

                    if (r["CorreoElectronico"] != System.DBNull.Value)
                        c.correoElectronico = r["CorreoElectronico"].ToString();

                    if (r["IndicadorTipoDireccion"] != System.DBNull.Value)
                        c.indicadorTipoDireccion = r["IndicadorTipoDireccion"].ToString();

                    if (r["IndicativoDeDireccionExterna"] != System.DBNull.Value)
                        c.indicativoDeDireccionExterna = r["IndicativoDeDireccionExterna"].ToString();

                    if (r["Linea1DeLaDireccionExterna"] != System.DBNull.Value)
                        c.linea1DeLaDireccionExterna = r["Linea1DeLaDireccionExterna"].ToString();

                    if (r["Linea2DeLaDireccionExterna"] != System.DBNull.Value)
                        c.linea2DeLaDireccionExterna = r["Linea2DeLaDireccionExterna"].ToString();

                    if (r["Linea3DeLaDireccionExterna"] != System.DBNull.Value)
                        c.linea3DeLaDireccionExterna = r["Linea3DeLaDireccionExterna"].ToString();

                    if (r["Linea4DeLaDireccionExterna"] != System.DBNull.Value)
                        c.linea4DeLaDireccionExterna = r["Linea4DeLaDireccionExterna"].ToString();

                    if (r["Idioma"] != System.DBNull.Value)
                        c.idioma = r["Idioma"].ToString();

                    if (r["FechaActivacion"] != System.DBNull.Value)
                        c.fechaActivacion = Convert.ToDateTime(r["FechaActivacion"]);

                    if (r["CodContrato"] != System.DBNull.Value)
                        c.codContrato = r["CodContrato"].ToString();

                    if (r["TipoAutoconsumo"] != System.DBNull.Value)
                        c.tipoAutoconsumo = r["TipoAutoconsumo"].ToString();

                    if (r["TipoContratoATR"] != System.DBNull.Value)
                        c.tipoContratoATR = r["TipoContratoATR"].ToString();

                    if (r["TarifaATR"] != System.DBNull.Value)
                        c.tarifaATR = r["TarifaATR"].ToString();

                    if (r["PeriodicidadFacturacion"] != System.DBNull.Value)
                        c.periodicidadFacturacion = r["PeriodicidadFacturacion"].ToString();

                    if (r["TipodeTelegestion"] != System.DBNull.Value)
                        c.tipodeTelegestion = r["TipodeTelegestion"].ToString();

                    if (r["PotenciaPeriodo1"] != System.DBNull.Value)
                        c.potenciaPeriodo[1] = Convert.ToDouble(r["PotenciaPeriodo1"]);
                    if (r["PotenciaPeriodo2"] != System.DBNull.Value)
                        c.potenciaPeriodo[2] = Convert.ToDouble(r["PotenciaPeriodo2"]);
                    if (r["PotenciaPeriodo3"] != System.DBNull.Value)
                        c.potenciaPeriodo[3] = Convert.ToDouble(r["PotenciaPeriodo3"]);
                    if (r["PotenciaPeriodo4"] != System.DBNull.Value)
                        c.potenciaPeriodo[4] = Convert.ToDouble(r["PotenciaPeriodo4"]);
                    if (r["PotenciaPeriodo5"] != System.DBNull.Value)
                        c.potenciaPeriodo[5] = Convert.ToDouble(r["PotenciaPeriodo5"]);
                    if (r["PotenciaPeriodo6"] != System.DBNull.Value)
                        c.potenciaPeriodo[6] = Convert.ToDouble(r["PotenciaPeriodo6"]);

                    if (r["MarcaMedidaConPerdidas"] != System.DBNull.Value)
                        c.marcaMedidaConPerdidas = r["MarcaMedidaConPerdidas"].ToString();

                    if (r["TensionDelSuministro"] != System.DBNull.Value)
                        c.tensionDelSuministro = Convert.ToInt32(r["TensionDelSuministro"]);


                    #endregion

                    EndesaEntity.contratacion.xxi.XML_Datos o;
                    if (!dic.TryGetValue(c.cups, out o))
                        dic.Add(c.cups, c);
                    else
                    {
                        if (o.fechaActivacion < c.fechaActivacion)
                            o = c;
                    }
                }
                db.CloseConnection();

                return dic;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "EEXXI - " + tipo_solicitud,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;


            }
        }

        //public void GeneraExcelComprobacion()
        //{

        //    string nombre_archivo = @"c:\Temp\CASOS_EEXXI_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".XLSX";
        //    ExportExcel(nombre_archivo);
        //    MessageBox.Show("Se ha generado el informe en:" + System.Environment.NewLine
        //        + nombre_archivo,
        //        "EEXXI",
        //        MessageBoxButtons.OK,
        //        MessageBoxIcon.Information);
        //}


        private string Accion(int caso)
        {
            string accion = "";
            switch (caso)
            {
                case 1:
                    accion = "Crear Contrato y Carta";
                    break;
                case 2:
                    accion = "Crear Contrato y Carta. Crear incidencia";
                    break;
                case 3:
                    accion = "Crear Contrato y Carta. Crear incidencia";
                    break;
                case 4:
                    accion = "Crear Contrato y Carta";
                    break;
                case 5:
                    accion = "Crear Contrato y Carta";
                    break;
                case 6:
                    accion = "Crear incidencia";
                    break;
                case 7:
                    accion = "No hacer nada";
                    break;
                case 8:
                    accion = "Crear Contrato y Carta. Informar a usuario";
                    break;
            }

            return accion;
        }


        private void CreaCasos()
        {
            contratacion.PuntosEnVigor psat = new PuntosEnVigor();
            informe = new List<EndesaEntity.contratacion.xxi.InformeCasos>();
            contratacion.eexxi.Casos caso = new Casos();            

            try
            {
                foreach (KeyValuePair<string, List<EndesaEntity.contratacion.Inventario_Detalle_Estados_Tabla>> p in inventarioDetalleEstados.dic_tmp)
                {

                    EndesaEntity.contratacion.xxi.InformeCasos c = new EndesaEntity.contratacion.xxi.InformeCasos();


                    c.descripcion = sc.TipoProceso(p.Value[0].codigoproceso, p.Value[0].codigopaso);
                    c.cups20 = p.Value[0].cups22.Substring(0, 20);
                    c.cups22 = p.Value[0].cups22;
                    c.codigo_solicitud = p.Value[0].codigosolicitud;
                    string psath;
                    
                    if (c.descripcion == "ALTA")
                    {
                        c.fecha_activacion = p.Value[0].fechaactivacion;
                        c.tipo_xml = p.Value[0].codigoproceso + "/" + p.Value[0].codigopaso;
                    }
                    else
                    {
                        EndesaEntity.contratacion.Inventario_Tabla oo = new EndesaEntity.contratacion.Inventario_Tabla();
                        //if (inventario.dic.TryGetValue(p.Value[0].cups22.Substring(0, 20), out oo))
                        if (inventario.dic.TryGetValue(p.Value[0].cups22, out oo))
                            c.fecha_activacion = oo.fecha_alta;

                        c.fecha_baja = p.Value[0].fechaactivacion;
                    }


                    List<EndesaEntity.contratacion.PS_AT_Tabla> o;
                    c.existe_ps = psat.dic_cups20.TryGetValue(c.cups20, out o);

                    if (p.Value.Count() > 1)
                    {
                        c.existe_baja = sc.TipoProceso(p.Value[1].codigoproceso, p.Value[1].codigopaso) == "BAJA";
                        if (c.existe_baja)
                        {
                            c.fecha_baja = p.Value[1].fechaactivacion;
                        }
                    }

                    if (c.existe_ps)
                    {
                        c.empresa = o[0].empresa;
                        c.estado_contrato_ps = o[0].estado_contrato_descripcion;
                    }
                    else
                    {
                        c.empresa = "N/A";
                        c.estado_contrato_ps = "N/A";
                    }

                    caso.GetCaso(c.descripcion, c.existe_baja, c.existe_ps, c.existe_ps ? o[0].empresa : null, c.existe_ps ? o[0].estado_contrato_id : 0);

                    c.caso = caso.caso;
                    c.acciones = caso.acciones;
                    c.crear_contrato = caso.crear_contrato;
                    c.crear_incidencia = caso.crear_incidencia;
                    c.realizar_seguimiento = caso.realizar_seguimiento;
                    informe.Add(c);

                }

                caso.GuardaCasos(informe);


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "EEXXI - CreaCasos",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }


        }



        private void ExportExcel(string fichero)
        {
            int f = 1;
            int c = 1;

            FileInfo fileInfo = new FileInfo(fichero);

            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("EEXXI");

            var headerCells = workSheet.Cells[1, 1, 1, 25];
            var headerFont = headerCells.Style.Font;

            headerFont.Bold = true;
            workSheet.Cells[f, c].Value = "CUPS"; c++;
            workSheet.Cells[f, c].Value = "FECHA ALTA"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "EXISTE ALTA"; c++;
            workSheet.Cells[f, c].Value = "EXISTE PS"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA PS"; c++;
            workSheet.Cells[f, c].Value = "ESTADO CONTRATO PS"; c++;
            workSheet.Cells[f, c].Value = "CREAR CONTRATO"; c++;
            workSheet.Cells[f, c].Value = "CREAR INCIDENCIA"; c++;
            workSheet.Cells[f, c].Value = "REALIZAR SEGUIMIENTO"; c++;
            workSheet.Cells[f, c].Value = "ACCIONES"; c++;


            for (int i = 0; i < informe.Count; i++)
            {
                c = 1;
                f++;
                workSheet.Cells[f, c].Value = informe[i].cups20; c++;
                workSheet.Cells[f, c].Value = informe[i].fecha_activacion; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                if (informe[i].fecha_baja > DateTime.MinValue)
                    workSheet.Cells[f, c].Value = informe[i].fecha_baja; workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = informe[i].existe_ps_hist; c++;
                //workSheet.Cells[f, c].Value = informe[i].caso; c++;
                //workSheet.Cells[f, c].Value = informe[i].tipo_xml; c++;
                //workSheet.Cells[f, c].Value = informe[i].descripcion; c++;
                //workSheet.Cells[f, c].Value = informe[i].existe_baja; c++;
                workSheet.Cells[f, c].Value = informe[i].existe_ps; c++;
                workSheet.Cells[f, c].Value = informe[i].empresa; c++;
                workSheet.Cells[f, c].Value = informe[i].estado_contrato_ps; c++;
                workSheet.Cells[f, c].Value = informe[i].crear_contrato; c++;
                workSheet.Cells[f, c].Value = informe[i].crear_incidencia; c++;
                workSheet.Cells[f, c].Value = informe[i].realizar_seguimiento; c++;
                workSheet.Cells[f, c].Value = informe[i].acciones; c++;
            }

            var allCells = workSheet.Cells[1, 1, f, c];
            allCells.AutoFitColumns();
            excelPackage.Save();

        }

        public Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> Carga_Entradas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>> d =
                new Dictionary<string, List<EndesaEntity.contratacion.xxi.XML_Min>>();

            try
            {
                strSql = "SELECT t.CUPS, t.CodigoDelProceso, t.CodigoDePaso, t.CodigoDeSolicitud,"
                    + " t.CodigoREEEmpresaEmisora, t.CodigoREEEmpresaDestino"
                    + " FROM eexxi_solicitudes_tmp t";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.XML_Min c = new EndesaEntity.contratacion.xxi.XML_Min();
                    c.proceso = r["CodigoDelProceso"].ToString();
                    c.paso = r["CodigoDePaso"].ToString();
                    c.solicitud = r["CodigoDeSolicitud"].ToString();
                    c.codigo_ree_empresa_emisora = r["CodigoREEEmpresaEmisora"].ToString();
                    c.codigo_ree_empresa_destino = r["CodigoREEEmpresaDestino"].ToString();

                    List<EndesaEntity.contratacion.xxi.XML_Min> o;
                    if (!d.TryGetValue(r["CUPS"].ToString(), out o))
                    {
                        o = new List<EndesaEntity.contratacion.xxi.XML_Min>();
                        o.Add(c);
                        d.Add(r["CUPS"].ToString(), o);
                    }
                    else
                        o.Add(c);
                }
                db.CloseConnection();

                strSql = "SELECT t.CUPS, t.CodigoDelProceso, t.CodigoDePaso, t.CodigoDeSolicitud,"
                    + " t.CodigoREEEmpresaEmisora, t.CodigoREEEmpresaDestino"
                    + " FROM eexxi_solicitudes t";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.XML_Min c = new EndesaEntity.contratacion.xxi.XML_Min();
                    c.proceso = r["CodigoDelProceso"].ToString();
                    c.paso = r["CodigoDePaso"].ToString();
                    c.solicitud = r["CodigoDeSolicitud"].ToString();
                    c.codigo_ree_empresa_emisora = r["CodigoREEEmpresaEmisora"].ToString();
                    c.codigo_ree_empresa_destino = r["CodigoREEEmpresaDestino"].ToString();

                    List<EndesaEntity.contratacion.xxi.XML_Min> o;
                    if (!d.TryGetValue(r["CUPS"].ToString(), out o))
                    {
                        o = new List<EndesaEntity.contratacion.xxi.XML_Min>();
                        o.Add(c);
                        d.Add(r["CUPS"].ToString(), o);
                    }
                    else
                        o.Add(c);
                }
                db.CloseConnection();

                strSql = "SELECT t.CUPS, t.CodigoDelProceso, t.CodigoDePaso, t.CodigoDeSolicitud,"
                    + " t.CodigoREEEmpresaEmisora, t.CodigoREEEmpresaDestino"
                    + " FROM eexxi_solicitudes_t101 t";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.XML_Min c = new EndesaEntity.contratacion.xxi.XML_Min();
                    c.proceso = r["CodigoDelProceso"].ToString();
                    c.paso = r["CodigoDePaso"].ToString();
                    c.solicitud = r["CodigoDeSolicitud"].ToString();
                    c.codigo_ree_empresa_emisora = r["CodigoREEEmpresaEmisora"].ToString();
                    c.codigo_ree_empresa_destino = r["CodigoREEEmpresaDestino"].ToString();

                    List<EndesaEntity.contratacion.xxi.XML_Min> o;
                    if (!d.TryGetValue(r["CUPS"].ToString(), out o))
                    {
                        o = new List<EndesaEntity.contratacion.xxi.XML_Min>();
                        o.Add(c);
                        d.Add(r["CUPS"].ToString(), o);
                    }
                    else
                        o.Add(c);
                }
                db.CloseConnection();

                return d;

            }
            catch(Exception ex)
            {
                return null;
            }
            
        }

        public void GuardadoBBDD(List<EndesaEntity.contratacion.xxi.XML_Datos> lista, string tabla)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";
            bool firstOnly = true;
            int num_reg = 0;

            foreach (EndesaEntity.contratacion.xxi.XML_Datos xml in lista)
            {
                if (firstOnly)
                {
                    strSql = "replace into " + tabla + " (CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,"
                        + " CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud, CUPS, CNAE, PotenciaExtension,"
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

                if (xml.fechaSolicitud != null)
                    strSql += ", '" + xml.fechaSolicitud.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                else
                    strSql += ", null";

                if (xml.cups != null)
                    strSql += ", '" + xml.cups + "'";
                else
                    strSql += ", null";

                if (xml.cnae != null)
                    strSql += ", " + xml.cnae;
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

                if (xml.fechaActivacion != null)
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
                db = new EndesaBusiness.servidores.MySQLDB(EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql.Substring(0, strSql.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                num_reg = 0;
                strSql = "";
            }


        }
               

        //public void XML_T105_Classic_To_Excel(List<EndesaEntity.contratacion.xxi.XML_Datos> lista_XML)
        //{

        //    int f = 0;
        //    int c = 0;
        //    int d = 0;

        //    string fichero = "";

        //    fichero = param.GetValue("RutaSalidaExcelAltas", DateTime.Now, DateTime.Now)
        //         + param.GetValue("NombreArchivoExcelAltas", DateTime.Now, DateTime.Now)
        //          + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
        //          + ".xlsx";

        //    FileInfo fileInfo = new FileInfo(fichero);

        //    if (fileInfo.Exists)
        //        fileInfo.Delete();

        //    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
        //    ExcelPackage excelPackage = new ExcelPackage(fileInfo);
        //    var workSheet = excelPackage.Workbook.Worksheets.Add("T105");

        //    var headerCells = workSheet.Cells[1, 1, 1, 28];
        //    var headerFont = headerCells.Style.Font;
        //    c = 0;

        //    //headerFont.Bold = true;

        //    foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista_XML)
        //    {
        //        c++;
        //        d = c + 1;
        //        f = 1;

        //        #region Cabecera
        //        workSheet.Cells[f, c].Value = "CodigoREEEmpresaEmisora";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.codigoREEEmpresaEmisora; f++;

        //        workSheet.Cells[f, c].Value = "CodigoREEEmpresaDestino";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.codigoREEEmpresaDestino; f++;

        //        workSheet.Cells[f, c].Value = "CodigoDelProceso";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.codigoDelProceso; f++;

        //        workSheet.Cells[f, c].Value = "CodigoDePaso";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.codigoDePaso; f++;

        //        workSheet.Cells[f, c].Value = "CodigoDeSolicitud";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.codigoDeSolicitud; f++;

        //        workSheet.Cells[f, c].Value = "FechaSolicitud";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.fechaSolicitud.ToString("dd/MM/yyyy"); f++;

        //        workSheet.Cells[f, c].Value = "CUPS";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.cups; f++;

        //        #endregion

        //        #region RegistrodePuntoSuministro
        //        workSheet.Cells[f, c].Value = "CNAE";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.cnae; f++;

        //        workSheet.Cells[f, c].Value = "PotenciaExtension";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.potenciaExtension; f++;

        //        workSheet.Cells[f, c].Value = "PotenciaDeAcceso";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.potenciaDeAcceso; f++;

        //        workSheet.Cells[f, c].Value = "PotenciaInstAT";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));

        //        workSheet.Cells[f, d].Value = p.potenciaInstAT; f++;

        //        workSheet.Cells[f, c].Value = "IndicativoDeInterrumpibilidad";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(192, 192, 192));
        //        workSheet.Cells[f, d].Value = p.indicativoDeInterrumpibilidad; f++;

        //        workSheet.Cells[f, c].Value = "DireccionPS";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0)); f++;


        //        workSheet.Cells[f, c].Value = "Pais";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.pais; f++;

        //        workSheet.Cells[f, c].Value = "Provincia";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.provincia; f++;

        //        workSheet.Cells[f, c].Value = "Municipio";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.municipio; f++;

        //        workSheet.Cells[f, c].Value = "Poblacion";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.poblacion; f++;

        //        workSheet.Cells[f, c].Value = "DescripcionPoblacion";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.descripcionPoblacion; f++;

        //        workSheet.Cells[f, c].Value = "TipoVia";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.tipoVia; f++;

        //        workSheet.Cells[f, c].Value = "CodPostal";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.codPostal; f++;

        //        workSheet.Cells[f, c].Value = "Calle";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.calle; f++;

        //        workSheet.Cells[f, c].Value = "NumeroFinca";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.numeroFinca; f++;

        //        workSheet.Cells[f, c].Value = "Piso";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));

        //        workSheet.Cells[f, d].Value = p.piso; f++;

        //        workSheet.Cells[f, c].Value = "TipoAclaradorFinca";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.tipoAclaradorFinca; f++;


        //        workSheet.Cells[f, c].Value = "AclaradorFinca";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.aclaradorFinca; f++;

        //        #endregion

        //        #region Cliente
        //        workSheet.Cells[f, c].Value = "TipoIdentificador";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.identificador; f++;

        //        workSheet.Cells[f, c].Value = "Identificador";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.; f++;

        //        workSheet.Cells[f, c].Value = "TipoPersona";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.cliente_TipoPersona; f++;

        //        workSheet.Cells[f, c].Value = "RazonSocial";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.razonSocial; f++;

        //        workSheet.Cells[f, c].Value = "PrefijoPais";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.cliente_PrefijoPais; f++;

        //        workSheet.Cells[f, c].Value = "Telefono";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.cliente_Numero; f++;

        //        workSheet.Cells[f, c].Value = "IndicadorTipoDireccion";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.cliente_IndicadorTipoDireccion; f++;

        //        workSheet.Cells[f, c].Value = "Linea1DeLaDireccionExterna";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.linea1DeLaDireccionExterna; f++;

        //        workSheet.Cells[f, c].Value = "Linea2DeLaDireccionExterna";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.linea2DeLaDireccionExterna; f++;

        //        workSheet.Cells[f, c].Value = "Linea3DeLaDireccionExterna";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.linea3DeLaDireccionExterna; f++;

        //        workSheet.Cells[f, c].Value = "Linea4DeLaDireccionExterna";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.linea4DeLaDireccionExterna; f++;

        //        workSheet.Cells[f, c].Value = "DireccionPS";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0)); f++;


        //        for (int x = 1; x <= p.potenciaPeriodo.Length; x++)
        //        {
        //            if (p.potenciaPeriodo[x] > 0)
        //            {
        //                workSheet.Cells[f, c].Value = "Potencia Periodo" + x;
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.potenciaPeriodo[x]; f++;
        //            }

        //        }

        //        workSheet.Cells[f, c].Value = "MarcaMedidaConPerdidas";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.marcaMedidaConPerdidas; f++;

        //        workSheet.Cells[f, c].Value = "TensionDelSuministro";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //        workSheet.Cells[f, d].Value = p.tensionDelSuministro; f++;


        //        workSheet.Cells[f, c].Value = "PuntosDeMedida";
        //        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0)); f++;

        //        for (int x = 0; x < p.lista_PM.Count(); x++)
        //        {
        //            workSheet.Cells[f, c].Value = "CodPM";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].codPM; f++;

        //            workSheet.Cells[f, c].Value = "TipoMovimiento";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].tipoMovimiento; f++;

        //            workSheet.Cells[f, c].Value = "TipoPM";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].tipoPM; f++;

        //            workSheet.Cells[f, c].Value = "ModoLectura";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].modoLectura; f++;

        //            workSheet.Cells[f, c].Value = "Funcion";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].funcion; f++;

        //            workSheet.Cells[f, c].Value = "TelefonoTelemedida";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].telefonoTelemedida; f++;

        //            workSheet.Cells[f, c].Value = "TensionPM";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].tensionPM; f++;

        //            workSheet.Cells[f, c].Value = "FechaVigor";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].fechaVigor.ToString("yyyy-MM-dd"); f++;

        //            workSheet.Cells[f, c].Value = "FechaAlta";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].fechaAlta.ToString("yyyy-MM-dd"); f++;

        //            workSheet.Cells[f, c].Value = "FechaBaja";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //            workSheet.Cells[f, d].Value = p.lista_PM[x].fechaBaja.ToString("yyyy-MM-dd"); f++;

        //            workSheet.Cells[f, c].Value = "Aparatos";
        //            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0)); f++;

        //            for (int y = 0; y < p.lista_PM[x].lista_aparatos.Count; y++)
        //            {
        //                workSheet.Cells[f, c].Value = "TipoAparato";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].tipoAparato; f++;

        //                workSheet.Cells[f, c].Value = "MarcaAparato";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].marcaAparato; f++;

        //                workSheet.Cells[f, c].Value = "ModeloMarca";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].modeloMarca; f++;

        //                workSheet.Cells[f, c].Value = "TipoMovimiento";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].tipoMovimiento; f++;

        //                workSheet.Cells[f, c].Value = "TipoEquipoMedida";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].tipoEquipoMedida; f++;

        //                workSheet.Cells[f, c].Value = "TipoPropiedadAparato";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].tipoPropiedadAparato; f++;

        //                workSheet.Cells[f, c].Value = "TipoDHEdM";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].tipoDHEdM; f++;

        //                workSheet.Cells[f, c].Value = "ModoMedidaPotencia";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].modoMedidaPotencia; f++;

        //                workSheet.Cells[f, c].Value = "CodPrecinto";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].codPrecinto; f++;

        //                workSheet.Cells[f, c].Value = "PeriodoFabricacion";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].periodoFabricacion; f++;

        //                workSheet.Cells[f, c].Value = "NumeroSerie";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].numeroSerie; f++;

        //                workSheet.Cells[f, c].Value = "FuncionAparato";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].funcionAparato; f++;

        //                workSheet.Cells[f, c].Value = "NumIntegradores";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].numIntegradores; f++;

        //                workSheet.Cells[f, c].Value = "ConstanteEnergia";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].constanteEnergia; f++;

        //                workSheet.Cells[f, c].Value = "ConstanteMaximetro";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].constanteMaximetro; f++;

        //                workSheet.Cells[f, c].Value = "RuedasEnteras";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].ruedasEnteras; f++;

        //                workSheet.Cells[f, c].Value = "RuedasDecimales";
        //                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(128, 64, 0));
        //                workSheet.Cells[f, d].Value = p.lista_PM[x].lista_aparatos[y].ruedasDecimales; f++;

        //            }
        //        }

        //        #endregion



        //    }

        //}

        public List<EndesaEntity.contratacion.xxi.XML_Datos> Completa_T105_con_T101(List<EndesaEntity.contratacion.xxi.XML_Datos> lista_T105)
        {
            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic_t101 =
                new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();
            // List<EndesaEntity.contratacion.xxi.XML_Datos> resultado_lista_T101 = new List<EndesaEntity.contratacion.xxi.XML_Datos>();
            //List<string> lista_codigos_soliditud = new List<string>();
            Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> dic_codigos_solicitud =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud>();

            T101 sol_t101 = new T101();

            try
            {

                // Capturamos todos los código de solicitud de la tabla eexxi_solicitudes_t101
                foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista_T105)
                {
                    EndesaEntity.contratacion.xxi.Cups_Solicitud c = new EndesaEntity.contratacion.xxi.Cups_Solicitud();
                    c.cups = p.cups;
                    c.solicitud = p.codigoDeSolicitud;

                    EndesaEntity.contratacion.xxi.Cups_Solicitud o;
                    if (!dic_codigos_solicitud.TryGetValue(c.solicitud + "_" + c.cups, out o))
                        dic_codigos_solicitud.Add(c.solicitud + "_" + c.cups, c);
                }
                    
                    

                //dic_t101 = BuscaDatosSolicitudXML("eexxi_solicitudes_t101", "T1", "01", lista_codigos_soliditud);
                //dic_t101 = BuscaDatosSolicitudXML("eexxi_solicitudes_t101", "T1", "01", dic_codigos_solicitud);

                // Completamos los datos del T105 con los datos recogidos de los T101 correspondientes

                foreach (EndesaEntity.contratacion.xxi.XML_Datos p in lista_T105)
                {
                    EndesaEntity.contratacion.xxi.XML_Datos t101 =
                        sol_t101.GetSolicitud_Cups(p.codigoDeSolicitud, p.cups);

                    if(t101 != null)
                    {

                        #region Cliente
                        p.paisCliente = t101.paisCliente;
                        p.provinciaCliente = t101.provinciaCliente;
                        p.municipioCliente = t101.municipioCliente;
                        p.codPostalCliente = t101.codPostalCliente; 
                        p.tipoViaCliente = t101.tipoViaCliente;
                        p.calleCliente = t101.calleCliente;
                        p.numeroFincaCliente = t101.numeroFincaCliente;
                        p.numero = t101.numero;

                        #endregion

                        p.codigoREEEmpresaEmisora = t101.codigoREEEmpresaEmisora;
                        p.codigoREEEmpresaDestino = t101.codigoREEEmpresaDestino;                        
                        p.fechaSolicitud = t101.fechaSolicitud;
                        p.cups = t101.cups;
                        p.cnae = t101.cnae;

                        #region DireccionPS
                        p.pais = t101.pais;
                        p.provincia = t101.provincia;
                        p.municipio = t101.municipio;
                        p.poblacion = t101.poblacion;
                        //p.descripcionPoblacion
                        //p.tipoVia ya no se utiliza
                        p.codPostal = t101.codPostal;
                        p.calle = t101.calle;
                        p.numeroFinca = t101.numeroFinca;
                        p.aclaradorFinca = t101.aclaradorFinca;

                        #endregion
                                               

                        p.tipoIdentificador = t101.tipoIdentificador;
                        p.identificador = t101.identificador;
                        p.tipoPersona = t101.tipoPersona;
                        p.razonSocial = t101.razonSocial;

                        p.prefijoPais = t101.prefijoPais;
                        p.numero = t101.numero;

                        p.indicadorTipoDireccion = t101.indicadorTipoDireccion;

                        p.indicativoDeDireccionExterna = t101.indicativoDeDireccionExterna;
                        //p.linea1DeLaDireccionExterna = t101.calleCliente + " " + t101.numeroFincaCliente;
                        //p.linea2DeLaDireccionExterna = municipio.DesMunicipio(t101.municipioCliente);
                        //p.linea3DeLaDireccionExterna = t101.codPostalCliente + ", "
                        //    + provincias.DesProvincia(t101.codPostalCliente);
                        //p.linea4DeLaDireccionExterna = t101.linea4DeLaDireccionExterna;

                        p.linea1DeLaDireccionExterna = t101.linea1DeLaDireccionExterna;
                        p.linea2DeLaDireccionExterna = t101.linea2DeLaDireccionExterna;
                        p.linea3DeLaDireccionExterna = t101.linea3DeLaDireccionExterna;
                        p.linea4DeLaDireccionExterna = t101.linea4DeLaDireccionExterna;

                        p.idioma = t101.idioma;

                        p.codContrato = t101.codContrato;
                        p.tarifaATR = t101.tarifaATR;


                        for (int i = 1; i < 7; i++)
                        {
                            if (t101.potenciaPeriodo[i] > 0)
                            {
                                p.potenciaPeriodo[i] = t101.potenciaPeriodo[i];                               
                            }                            
                        }



                    }
                }

                return lista_T105;

            }
            catch(Exception e)
            {
                return null;
            }


        }

        
        

    }
}
