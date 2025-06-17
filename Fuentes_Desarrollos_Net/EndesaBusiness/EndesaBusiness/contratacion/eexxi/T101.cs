using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class T101
    {
        Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> dic;
        EndesaBusiness.global.Provincias provincias;
        EndesaBusiness.global.Municipios municipio;
        public T101()
        {
            provincias = new global.Provincias("eexxi_param_provincias", servidores.MySQLDB.Esquemas.CON);
            municipio = new global.Municipios();
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos> d
                = new Dictionary<string, EndesaEntity.contratacion.xxi.XML_Datos>();

            try
            {
                strSql = "SELECT sol.*, p.DesProvincia, p2.DesProvincia as DesProvinciaCliente "
                    + " FROM eexxi_solicitudes_t101 sol"
                    + " INNER JOIN eexxi_solicitudes_tmp t ON"
                    + " t.CodigoDeSolicitud = sol.CodigoDeSolicitud AND"
                    + " t.CUPS = sol.CUPS AND"
                    + " t.CodigoDelProceso = 'T1' AND"
                    + " t.CodigoDePaso = '05'"
                    + " INNER JOIN eexxi_param_codigos_tarifas_atr tarifa ON"
                    + " tarifa.tarifa_atr = sol.TarifaATR"
                    + " LEFT OUTER JOIN eexxi_param_provincias p ON"
                    + " substr(p.CodigoPostal,1,2) = sol.Provincia"
                    + " LEFT OUTER JOIN eexxi_param_provincias p2 ON"
                    + " substr(p2.CodigoPostal,1,2) = sol.ProvinciaCliente"
                    + " WHERE"
                    + " tarifa.eexxi = 'S' and"
                    + " sol.PotenciaPeriodo6 > 50000";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    #region Campos

                    EndesaEntity.contratacion.xxi.XML_Datos c = new EndesaEntity.contratacion.xxi.XML_Datos();
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

                    EndesaEntity.contratacion.xxi.XML_Datos o;
                    if (!d.TryGetValue(c.codigoDeSolicitud + "_" + c.cups, out o))                    
                        d.Add(c.codigoDeSolicitud + "_" + c.cups, c);
                    

                }
                db.CloseConnection();

                return d;

            }
            catch(Exception ex)
            {

                MessageBox.Show(ex.Message,
               "T101 - Carga",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
                return null;
            }
        }

        public EndesaEntity.contratacion.xxi.XML_Datos GetSolicitud_Cups(string codigoDeSolicitud, string cups)
        {
            EndesaEntity.contratacion.xxi.XML_Datos o;
            if (dic.TryGetValue(codigoDeSolicitud + "_" + cups, out o))
                return o;
            else
                return null;

        }
    }
}
