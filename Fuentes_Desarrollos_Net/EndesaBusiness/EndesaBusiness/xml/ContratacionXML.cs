using EndesaBusiness.contratacion.eexxi;
using EndesaBusiness.servidores;
using EndesaEntity.cnmc.V21_2019_12_17;
using EndesaEntity.xml;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.Xml;
using System.Xml.Schema;
using EndesaEntity.cnmc.V30_2022_21_01;

namespace EndesaBusiness.xml
{
    public class ContratacionXML
    {
        Dictionary<string, List<EndesaEntity.xml.CNMC_t_cabecera>> dic_cabeceras;
        List<string> lista_procesos;
        List<string> lista_pasos;
        List<EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301> lista_a301;
        utilidades.Param p;
        public ContratacionXML()
        {
            p = new utilidades.Param("cnmc_param", servidores.MySQLDB.Esquemas.CON);
            dic_cabeceras = Carga();

        }

        private Dictionary<string, List<EndesaEntity.xml.CNMC_t_cabecera>> Carga()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, List<EndesaEntity.xml.CNMC_t_cabecera>> d =
                new Dictionary<string, List<EndesaEntity.xml.CNMC_t_cabecera>>();
            try
            {
                strSql = "select id, CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,"
                    + " CodigoDePaso, CodigoDeSolicitud, FechaSolicitud, CUPS"
                    + " from cnmc_t_cabecera order by id";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.xml.CNMC_t_cabecera c = new EndesaEntity.xml.CNMC_t_cabecera();

                    c.id = Convert.ToInt32(r["id"]);

                    if (r["CodigoREEEmpresaEmisora"] != System.DBNull.Value)
                        c.CodigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();
                    if (r["CodigoREEEmpresaDestino"] != System.DBNull.Value)
                        c.CodigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();
                    if (r["CodigoDelProceso"] != System.DBNull.Value)
                       c.CodigoDelProceso = r["CodigoDelProceso"].ToString();
                    //irh
                    var reemplazos = new Dictionary<string, string>
                    {
                          { "A3", "AD" },
                          { "C1", "ACC" },
                          { "B1", "BA" },
                          { "M1", "MOD" },
                          { "C2", "ACCMOD" }
                     };

                    if (reemplazos.ContainsKey(c.CodigoDelProceso))
                    {
                        c.CodigoDelProceso = reemplazos[c.CodigoDelProceso];
                    }

                    if (r["CodigoDePaso"] != System.DBNull.Value)
                        c.CodigoDePaso = r["CodigoDePaso"].ToString();

                    if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                        c.CodigoDeSolicitud = r["CodigoDeSolicitud"].ToString();

                    if (r["FechaSolicitud"] != System.DBNull.Value)
                        c.FechaSolicitud = Convert.ToDateTime(r["FechaSolicitud"]); 

                    if (r["CUPS"] != System.DBNull.Value)
                        c.CUPS = r["CUPS"].ToString();



                    List<EndesaEntity.xml.CNMC_t_cabecera> o;
                    if(!d.TryGetValue(c.CodigoDelProceso + c.CodigoDePaso, out o))
                    {
                        o = new List<EndesaEntity.xml.CNMC_t_cabecera>();
                        o.Add(c);
                        d.Add(c.CodigoDelProceso + c.CodigoDePaso, o);
                    }
                    else
                        o.Add(c);


                }
                db.CloseConnection();

                return d;

            }catch(Exception e)
            {
                return null; ;
            }
        }

        public List<EndesaEntity.xml.CNMC_t_cabecera> GetLista()
        {
            dic_cabeceras = Carga();

            List<EndesaEntity.xml.CNMC_t_cabecera> d = new List<EndesaEntity.xml.CNMC_t_cabecera>();
            foreach (KeyValuePair<string, List<EndesaEntity.xml.CNMC_t_cabecera>> p in dic_cabeceras)
                foreach (EndesaEntity.xml.CNMC_t_cabecera pp in p.Value)
                    d.Add(pp);

            return d;
        }

        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 CargaDatos(int id)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml =
                new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301();

            try
            {

                strSql = "SELECT cab.CodigoREEEmpresaEmisora, cab.CodigoREEEmpresaDestino, cab.CodigoDelProceso,"
                    + " cab.CodigoDePaso, cab.CodigoDeSolicitud, cab.SecuencialDeSolicitud, cab.FechaSolicitud,"
                    + " cab.CUPS,sol.CNAE, sol.IndActivacion, sol.SolicitudTension, sol.TipoAutoconsumo, sol.TipoContratoATR,"
                    + " sol.TarifaATR,sol.Potencia1, sol.Potencia2, sol.Potencia3, sol.Potencia4, sol.Potencia5,"
                    + " sol.Potencia6, sol.ModoControlPotencia,con.PersonaDeContacto, con.PrefijoPais, con.Numero,"
                    + " cli.TipoIdentificador, cli.Identificador, cli.TipoPersona,cli.NombreDePila, cli.PrimerApellido,"
                    + " cli.SegundoApellido, cli.RazonSocial"
                    + " FROM cnmc_t_cabecera cab"
                    + " INNER JOIN cnmc_t_datossolicitud sol ON"
                    + " sol.id = cab.id"
                    + " INNER JOIN cnmc_t_cliente cli ON"
                    + " cli.id = cab.id"
                    + " LEFT JOIN cnmc_t_contacto con ON"
                    + " con.id = cab.id"
                    + " WHERE cab.id = " + id;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    xml.Cabecera.CodigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();
                    xml.Cabecera.CodigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();
                    xml.Cabecera.CodigoDelProceso = r["CodigoDelProceso"].ToString();
                    xml.Cabecera.CodigoDePaso = r["CodigoDePaso"].ToString();
                    xml.Cabecera.CodigoDeSolicitud = r["CodigoDeSolicitud"].ToString();
                    xml.Cabecera.SecuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();
                    xml.Cabecera.FechaSolicitud = r["FechaSolicitud"].ToString();
                    xml.Cabecera.CUPS = r["CUPS"].ToString();

                    if (r["CNAE"] != System.DBNull.Value)
                        xml.Alta.DatosSolicitud.cnae = r["CNAE"].ToString();

                    if (r["IndActivacion"] != System.DBNull.Value)
                        xml.Alta.DatosSolicitud.IndActivacion = r["IndActivacion"].ToString();

                    if (r["SolicitudTension"] != System.DBNull.Value)
                        xml.Alta.DatosSolicitud.SolicitudTension = r["SolicitudTension"].ToString();
                    else
                        xml.Alta.DatosSolicitud.SolicitudTension = "N";

                    if (r["TipoAutoconsumo"] != System.DBNull.Value)
                        xml.Alta.Contrato.TipoAutoconsumo = r["TipoAutoconsumo"].ToString();

                    if (r["TipoContratoATR"] != System.DBNull.Value)
                        xml.Alta.Contrato.TipoContratoATR = r["TipoContratoATR"].ToString();

                    if (r["TarifaATR"] != System.DBNull.Value)
                        xml.Alta.Contrato.CondicionesContractuales.TarifaATR = r["TarifaATR"].ToString();

                    for (int j = 1; j <= 6; j++)
                        if (r["Potencia" + j] != System.DBNull.Value)
                        {
                            EndesaEntity.cnmc.V21_2019_12_17.Potencia p = new Potencia();
                            p.periodo = j.ToString();
                            p.potencia = r["Potencia" + j].ToString();

                            xml.Alta.Contrato.CondicionesContractuales.
                               PotenciasContratadas.Potencia.Add(p);


                        }
                           


                }
                db.CloseConnection();
                return xml;
            }
            catch(Exception ex)
            {
                
                return null;
            }

            
        }

        //public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 CargaDatosV30(int id)
        //{
        //    servidores.MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql;

        //    EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml =
        //        new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301();

        //    try
        //    {

        //        strSql = "SELECT cab.CodigoREEEmpresaEmisora, cab.CodigoREEEmpresaDestino, cab.CodigoDelProceso,"
        //            + " cab.CodigoDePaso, cab.CodigoDeSolicitud, cab.SecuencialDeSolicitud, cab.FechaSolicitud,"
        //            + " cab.CUPS,sol.CNAE, sol.IndActivacion, sol.SolicitudTension, sol.TipoAutoconsumo, sol.TipoContratoATR,"
        //            + " sol.TarifaATR,sol.Potencia1, sol.Potencia2, sol.Potencia3, sol.Potencia4, sol.Potencia5,"
        //            + " sol.Potencia6, sol.ModoControlPotencia,con.PersonaDeContacto, con.PrefijoPais, con.Numero,"
        //            + " cli.TipoIdentificador, cli.Identificador, cli.TipoPersona,cli.NombreDePila, cli.PrimerApellido,"
        //            + " cli.SegundoApellido, cli.RazonSocial"
        //            + " FROM cnmc_t_cabecera cab"
        //            + " INNER JOIN cnmc_t_datossolicitud sol ON"
        //            + " sol.id = cab.id"
        //            + " INNER JOIN cnmc_t_cliente cli ON"
        //            + " cli.id = cab.id"
        //            + " LEFT JOIN cnmc_t_contacto con ON"
        //            + " con.id = cab.id"
        //            + " WHERE cab.id = " + id;
        //        db = new MySQLDB(MySQLDB.Esquemas.CON);
        //        command = new MySqlCommand(strSql, db.con);
        //        r = command.ExecuteReader();
        //        while (r.Read())
        //        {
        //            xml.Cabecera.CodigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();
        //            xml.Cabecera.CodigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();
        //            xml.Cabecera.CodigoDelProceso = r["CodigoDelProceso"].ToString();

        //            xml.Cabecera.CodigoDePaso = r["CodigoDePaso"].ToString();
        //            xml.Cabecera.CodigoDeSolicitud = r["CodigoDeSolicitud"].ToString();
        //            xml.Cabecera.SecuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();
        //            xml.Cabecera.FechaSolicitud = r["FechaSolicitud"].ToString();
        //            xml.Cabecera.CUPS = r["CUPS"].ToString();

        //            if (r["CNAE"] != System.DBNull.Value)
        //                xml.Alta.DatosSolicitud.cnae = r["CNAE"].ToString();

        //            if (r["IndActivacion"] != System.DBNull.Value)
        //                xml.Alta.DatosSolicitud.IndActivacion = r["IndActivacion"].ToString();

        //            if (r["SolicitudTension"] != System.DBNull.Value)
        //                xml.Alta.DatosSolicitud.SolicitudTension = r["SolicitudTension"].ToString();
        //            else
        //                xml.Alta.DatosSolicitud.SolicitudTension = "N";

        //            if (r["TipoAutoconsumo"] != System.DBNull.Value)
        //            {
        //                xml.Alta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();
        //                xml.Alta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo = r["TipoAutoconsumo"].ToString();

        //                if (xml.Alta == null)
        //                    xml.Alta = new AltaA301();
        //                if (xml.Alta.Contrato == null)
        //                    xml.Alta.Contrato = new ContratoAlta();

        //                xml.Alta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();
        //                xml.Alta.Contrato.Autoconsumo.DatosCAU = new DatosCAUAlta();

        //                xml.Alta.Contrato.Autoconsumo.DatosCAU.TipoAutoconsumo = r["TipoAutoconsumo"].ToString();
        //            }


        //            if (r["TipoContratoATR"] != System.DBNull.Value)
        //                xml.Alta.Contrato.TipoContratoATR = r["TipoContratoATR"].ToString();

        //            if (r["TarifaATR"] != System.DBNull.Value)
        //                xml.Alta.Contrato.CondicionesContractuales.TarifaATR = r["TarifaATR"].ToString();

        //            for (int j = 1; j <= 6; j++)
        //                if (r["Potencia" + j] != System.DBNull.Value)
        //                {
        //                    EndesaEntity.cnmc.V21_2019_12_17.Potencia p = new Potencia();
        //                    p.periodo = j.ToString();
        //                    p.potencia = r["Potencia" + j].ToString();

        //                    xml.Alta.Contrato.CondicionesContractuales.
        //                       PotenciasContratadas.Potencia.Add(p);


        //                }



        //        }
        //        db.CloseConnection();
        //        return xml;
        //    }
        //    catch (Exception ex)
        //    {

        //        return null;
        //    }
        // }
        public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 CargaDatosV30(int id)
        {
            servidores.MySQLDB db  = null;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml =
                new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301();

            try
            {
                strSql = "SELECT cab.CodigoREEEmpresaEmisora, cab.CodigoREEEmpresaDestino, cab.CodigoDelProceso,"
                    + " cab.CodigoDePaso, cab.CodigoDeSolicitud, cab.SecuencialDeSolicitud, cab.FechaSolicitud,"
                    + " cab.CUPS, sol.CNAE, sol.IndActivacion, sol.SolicitudTension, sol.TipoAutoconsumo, sol.TipoContratoATR,"
                    + " sol.TarifaATR, sol.Potencia1, sol.Potencia2, sol.Potencia3, sol.Potencia4, sol.Potencia5,"
                    + " sol.Potencia6, sol.ModoControlPotencia, sol.TipoCUPS, sol.cau, sol.TipoSubseccion,"
                    + " sol.Colectivo, sol.PotInstaladaGen, sol.TipoInstalacion, sol.SSAA, sol.UnicoContrato,"
                    + " sol.RefCatastro, sol.CIL, sol.TecGenerado, sol.IndEsencial, sol.fechaPrevistaAccion,"
                    + " sol.TensionSolicitada, sol.CodContrato, sol.Fecha_finalizacion, sol.TipoActivacionPrevista,"
                    + " sol.CUPSPrincipal, sol.FechaActivacionPrevista, sol.TensionDelSuministro, sol.VAsTrafo,"
                    + " sol.PeriodicidadFacturacion, sol.TipoTelegestion,"
                    + " cli.TipoIdentificador, cli.Identificador, cli.TipoPersona, cli.NombreDePila,"
                    + " cli.PrimerApellido, cli.SegundoApellido, cli.RazonSocial"
                    + " FROM cnmc_t_cabecera cab"
                    + " INNER JOIN cnmc_t_datossolicitudnew sol ON sol.id = cab.id"
                    + " INNER JOIN cnmc_t_cliente cli ON cli.id = cab.id"
                    + " LEFT JOIN cnmc_t_contacto con ON con.id = cab.id"
                    + " WHERE cab.id = " + id;

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    xml.Cabecera.CodigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();
                    xml.Cabecera.CodigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();
                    xml.Cabecera.CodigoDelProceso = r["CodigoDelProceso"].ToString();
                    xml.Cabecera.CodigoDePaso = r["CodigoDePaso"].ToString();
                    xml.Cabecera.CodigoDeSolicitud = r["CodigoDeSolicitud"].ToString();
                    xml.Cabecera.SecuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();
                    xml.Cabecera.FechaSolicitud = r["FechaSolicitud"].ToString();
                    xml.Cabecera.CUPS = r["CUPS"].ToString();

                    var datos = xml.Alta.DatosSolicitud;
                    var contrato = xml.Alta.Contrato;

                    if (r["CNAE"] != DBNull.Value)
                        datos.cnae = r["CNAE"].ToString();

                    if (r["IndActivacion"] != DBNull.Value)
                        datos.IndActivacion = r["IndActivacion"].ToString();

                    if (r["SolicitudTension"] != DBNull.Value)
                        datos.SolicitudTension = r["SolicitudTension"].ToString();
                    else
                        datos.SolicitudTension = "N";

                    // AUTOCONSUMO
                    if (r["TipoAutoconsumo"] != DBNull.Value)
                    {
                        if (xml.Alta.Contrato.Autoconsumo == null)
                        {
                            xml.Alta.Contrato.Autoconsumo = new AutoconsumoSolicitudAlta();
                            xml.Alta.Contrato.Autoconsumo.DatosCAU = new DatosCAUAlta();
                            xml.Alta.Contrato.Autoconsumo.DatosSuministro = new DatosSuministroSolicitud();

                        }

                        var auto = xml.Alta.Contrato.Autoconsumo;
                        auto.DatosCAU.TipoAutoconsumo = r["TipoAutoconsumo"].ToString();

                        // NUEVOS CAMPOS de Autoconsumo
                        if (r["TipoCUPS"] != DBNull.Value)
                            auto.DatosSuministro.TipoCUPS = r["TipoCUPS"].ToString();

                        if (r["cau"] != DBNull.Value)
                            auto.DatosCAU.CAU = r["cau"].ToString();

                        if (r["TipoSubseccion"] != DBNull.Value)
                            auto.DatosCAU.TipoSubseccion = r["TipoSubseccion"].ToString();

                        if (r["Colectivo"] != DBNull.Value)
                            auto.DatosCAU.Colectivo = r["Colectivo"].ToString();

                        if (auto.DatosCAU.DatosInstGen == null)
                            xml.Alta.Contrato.Autoconsumo.DatosCAU = new DatosCAUAlta();

                        //xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen = auto.DatosCAU.DatosInstGen.


                        var datosInstGen = auto.DatosCAU.DatosInstGen;

                        if (r["PotInstaladaGen"] != DBNull.Value)
                            datosInstGen.PotInstaladaGen = r["PotInstaladaGen"].ToString();

                        if (r["TipoInstalacion"] != DBNull.Value)
                            datosInstGen.TipoInstalacion = r["TipoInstalacion"].ToString();

                        if (r["SSAA"] != DBNull.Value)
                            datosInstGen.SSAA = r["SSAA"].ToString();

                        if (r["UnicoContrato"] != DBNull.Value)
                            datosInstGen.UnicoContrato = r["UnicoContrato"].ToString();

                        if (r["RefCatastro"] != DBNull.Value)
                            datosInstGen.RefCatastro = r["RefCatastro"].ToString();

                        if (r["CIL"] != DBNull.Value)
                            datosInstGen.CIL = r["CIL"].ToString();

                        if (r["TecGenerado"] != DBNull.Value)
                            datosInstGen.TecGenerador = r["TecGenerado"].ToString();
                    }

                    if (r["TipoContratoATR"] != DBNull.Value)
                        contrato.TipoContratoATR = r["TipoContratoATR"].ToString();

                    if (r["TarifaATR"] != DBNull.Value)
                        contrato.CondicionesContractuales.TarifaATR = r["TarifaATR"].ToString();

                    // NUEVOS CAMPOS de DatosSolicitud
                    if (r["IndEsencial"] != DBNull.Value)
                        datos.IndEsencial = r["IndEsencial"].ToString();

                    if (r["fechaPrevistaAccion"] != DBNull.Value)
                        datos.fechaPrevistaAccion = r["fechaPrevistaAccion"].ToString();

                    if (r["TensionSolicitada"] != DBNull.Value)
                        datos.TensionSolicitada = r["TensionSolicitada"].ToString();

                    // Potencias
                    for (int j = 1; j <= 6; j++)
                    {
                        if (r["Potencia" + j] != DBNull.Value)
                        {
                            var p = new EndesaEntity.cnmc.V21_2019_12_17.Potencia
                            {
                                periodo = j.ToString(),
                                potencia = r["Potencia" + j].ToString()
                            };

                            contrato.CondicionesContractuales.PotenciasContratadas.Potencia.Add(p);
                        }
                    }

                    if (r["ModoControlPotencia"] != DBNull.Value)
                        contrato.CondicionesContractuales.ModoControlPotencia = r["ModoControlPotencia"].ToString();

                    // NUEVOS CAMPOS de Contrato
                    if (r["CodContrato"] != DBNull.Value)
                    {
                        if (contrato.IdContrato == null)
                          //  contrato.IdContrato = new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301.a                                
                               // .IdContrato();
                        contrato.IdContrato.CodContrato = r["CodContrato"].ToString();
                    }

                    if (r["Fecha_finalizacion"] != DBNull.Value)
                        contrato.FechaFinalizacion = r["Fecha_finalizacion"].ToString();

                    if (r["TipoActivacionPrevista"] != DBNull.Value)
                        contrato.TipoActivacionPrevista = r["TipoActivacionPrevista"].ToString();

                    if (r["CUPSPrincipal"] != DBNull.Value)
                        contrato.CUPSPrincipal = r["CUPSPrincipal"].ToString();

                    if (r["FechaActivacionPrevista"] != DBNull.Value)
                        contrato.FechaActivacionPrevista = r["FechaActivacionPrevista"].ToString();

                    if (r["TensionDelSuministro"] != DBNull.Value)
                        contrato.CondicionesContractuales.TensionDelSuministro = r["TensionDelSuministro"].ToString();

                    if (r["VAsTrafo"] != DBNull.Value)
                        contrato.CondicionesContractuales.VAsTrafo = r["VAsTrafo"].ToString();

                    if (r["PeriodicidadFacturacion"] != DBNull.Value)
                        contrato.CondicionesContractuales.PeriodicidadFacturacion = r["PeriodicidadFacturacion"].ToString();

                    if (r["TipoTelegestion"] != DBNull.Value)
                        contrato.CondicionesContractuales.TipoTelegestion = r["TipoTelegestion"].ToString();
                }

                db.CloseConnection();
                return xml;
            }
            catch (Exception ex)
            {
                db?.CloseConnection();
                return null;
            }
        }


        //public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 CargaDatosC1(int id)  // irh actualizar tabla cnmc_t_datossolicitudnew 
        //{
        //    servidores.MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader r;
        //    string strSql;

        //    EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml =
        //        new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101();

        //    try
        //    {

        //        strSql = "SELECT cab.CodigoREEEmpresaEmisora, cab.CodigoREEEmpresaDestino, cab.CodigoDelProceso,"
        //            + " cab.CodigoDePaso, cab.CodigoDeSolicitud, cab.SecuencialDeSolicitud, cab.FechaSolicitud,"
        //            + " cab.CUPS,sol.CNAE, sol.IndActivacion, sol.SolicitudTension, sol.TipoAutoconsumo, sol.TipoContratoATR,"
        //            + " sol.TarifaATR,sol.Potencia1, sol.Potencia2, sol.Potencia3, sol.Potencia4, sol.Potencia5,"
        //            + " sol.Potencia6, sol.ModoControlPotencia,con.PersonaDeContacto, con.PrefijoPais, con.Numero,"
        //            + " cli.TipoIdentificador, cli.Identificador, cli.TipoPersona,cli.NombreDePila, cli.PrimerApellido,"
        //            + " cli.SegundoApellido, cli.RazonSocial"
        //            + " FROM cnmc_t_cabecera cab"
        //            + " INNER JOIN cnmc_t_datossolicitud sol ON"
        //            + " sol.id = cab.id"
        //            + " INNER JOIN cnmc_t_cliente cli ON"
        //            + " cli.id = cab.id"
        //            + " LEFT JOIN cnmc_t_contacto con ON"
        //            + " con.id = cab.id"
        //            + " WHERE cab.id = " + id;
        //        db = new MySQLDB(MySQLDB.Esquemas.CON);
        //        command = new MySqlCommand(strSql, db.con);
        //        r = command.ExecuteReader();
        //        while (r.Read())
        //        {
        //            xml.Cabecera.CodigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();
        //            xml.Cabecera.CodigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();
        //            xml.Cabecera.CodigoDelProceso = r["CodigoDelProceso"].ToString();
        //            xml.Cabecera.CodigoDePaso = r["CodigoDePaso"].ToString();
        //            xml.Cabecera.CodigoDeSolicitud = r["CodigoDeSolicitud"].ToString();
        //            xml.Cabecera.SecuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();
        //            xml.Cabecera.FechaSolicitud = r["FechaSolicitud"].ToString();
        //            xml.Cabecera.CUPS = r["CUPS"].ToString();

        //            // DatosSolicitud_C1
        //            var datos = xml.CambiodeComercializadorSinCambios.DatosSolicitud;
        //            if (r["CNAE"] != DBNull.Value) datos.cnae = r["CNAE"].ToString();
        //            if (r["IndActivacion"] != DBNull.Value) datos.indActivacion = r["IndActivacion"].ToString();
        //            if (r["SolicitudTension"] != DBNull.Value) datos.solicitudTension = r["SolicitudTension"].ToString();
        //            if (r["TipoModificacion"] != DBNull.Value) datos.tipoModificacion = r["TipoModificacion"].ToString();
        //            if (r["TipoSolicitudAdministrativa"] != DBNull.Value) datos.tipoSolicitudAdministrativa = r["TipoSolicitudAdministrativa"].ToString();
        //            if (r["FechaPrevistaAccion"] != DBNull.Value) datos.fechaPrevistaAccion = r["FechaPrevistaAccion"].ToString();
        //            if (r["ContratacionIncondicionalPS"] != DBNull.Value) datos.contratacionIncondicionalPS = r["ContratacionIncondicionalPS"].ToString();
        //            if (r["ContratacionIncondicionalBS"] != DBNull.Value) datos.contratacionIncondicionalBS = r["ContratacionIncondicionalBS"].ToString();
        //            if (r["TensionSolicitada"] != DBNull.Value) datos.TensionSolicitada = r["TensionSolicitada"].ToString();

        //            // Cliente
        //            var cli = new EndesaEntity.cnmc.V21_2019_12_17.Cliente();
                                        
        //            if (r["TipoIdentificador"] != DBNull.Value) cli.IdCliente.TipoIdentificador = r["TipoIdentificador"].ToString();
        //            if (r["Identificador"] != DBNull.Value) cli.IdCliente.Identificador = r["Identificador"].ToString();
        //            if (r["TipoPersona"] != DBNull.Value) cli.IdCliente.TipoPersona = r["TipoPersona"].ToString();

        //            if (r["NombreDePila"] != DBNull.Value) cli.Nombre.NombreDePila = r["NombreDePila"].ToString();
        //            if (r["PrimerApellido"] != DBNull.Value) cli.Nombre.PrimerApellido = r["PrimerApellido"].ToString();
        //            if (r["SegundoApellido"] != DBNull.Value) cli.Nombre.SegundoApellido = r["SegundoApellido"].ToString();
        //            if (r["RazonSocial"] != DBNull.Value) cli.Nombre.RazonSocial = r["RazonSocial"].ToString();

        //            xml.CambiodeComercializadorSinCambios.Cliente = cli;

                  
        //        }
        //        db.CloseConnection();
        //        return xml;
        //    }
        //    catch (Exception ex)
        //    {

        //        return null;
        //    }


        //}
        public EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 CargaDatosC1(int id)  // irh actualizar tabla cnmc_t_datossolicitudnew 
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml =
                new EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101();

            try
            {

                strSql = "SELECT cab.CodigoREEEmpresaEmisora, cab.CodigoREEEmpresaDestino, cab.CodigoDelProceso,"
                    + " cab.CodigoDePaso, cab.CodigoDeSolicitud, cab.SecuencialDeSolicitud, cab.FechaSolicitud,"
                    + " cab.CUPS,sol.CNAE, sol.IndActivacion, sol.SolicitudTension, sol.TipoAutoconsumo, sol.TipoContratoATR,"
                    + " sol.TarifaATR,sol.Potencia1, sol.Potencia2, sol.Potencia3, sol.Potencia4, sol.Potencia5,"
                    + " sol.Potencia6, sol.ModoControlPotencia,con.PersonaDeContacto, con.PrefijoPais, con.Numero,"
                    + " cli.TipoIdentificador, cli.Identificador, cli.TipoPersona,cli.NombreDePila, cli.PrimerApellido,"
                    + " cli.SegundoApellido, cli.RazonSocial"
                    + " FROM cnmc_t_cabecera cab"
                    + " INNER JOIN cnmc_t_datossolicitudnew sol ON"
                    + " sol.id = cab.id"
                    + " INNER JOIN cnmc_t_cliente cli ON"
                    + " cli.id = cab.id"
                    + " LEFT JOIN cnmc_t_contacto con ON"
                    + " con.id = cab.id"
                    + " WHERE cab.id = " + id;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    xml.Cabecera.CodigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();
                    xml.Cabecera.CodigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();
                    xml.Cabecera.CodigoDelProceso = r["CodigoDelProceso"].ToString();
                    xml.Cabecera.CodigoDePaso = r["CodigoDePaso"].ToString();
                    xml.Cabecera.CodigoDeSolicitud = r["CodigoDeSolicitud"].ToString();
                    xml.Cabecera.SecuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();
                    xml.Cabecera.FechaSolicitud = r["FechaSolicitud"].ToString();
                    xml.Cabecera.CUPS = r["CUPS"].ToString();

                    // DatosSolicitud_C1
                    var datos = xml.CambiodeComercializadorSinCambios.DatosSolicitud;
                    if (r["CNAE"] != DBNull.Value) datos.cnae = r["CNAE"].ToString();
                    if (r["IndActivacion"] != DBNull.Value) datos.indActivacion = r["IndActivacion"].ToString();
                    if (r["SolicitudTension"] != DBNull.Value) datos.solicitudTension = r["SolicitudTension"].ToString();
                   // if (r["TipoModificacion"] != DBNull.Value) datos.tipoModificacion = r["TipoModificacion"].ToString();
                    //if (r["TipoSolicitudAdministrativa"] != DBNull.Value) datos.tipoSolicitudAdministrativa = r["TipoSolicitudAdministrativa"].ToString();
                   // if (r["FechaPrevistaAccion"] != DBNull.Value) datos.fechaPrevistaAccion = r["FechaPrevistaAccion"].ToString();
                  //  if (r["ContratacionIncondicionalPS"] != DBNull.Value) datos.contratacionIncondicionalPS = r["ContratacionIncondicionalPS"].ToString();
                   // if (r["ContratacionIncondicionalBS"] != DBNull.Value) datos.contratacionIncondicionalBS = r["ContratacionIncondicionalBS"].ToString();
                   // if (r["TensionSolicitada"] != DBNull.Value) datos.TensionSolicitada = r["TensionSolicitada"].ToString();

                    // Cliente
                    var cli = new EndesaEntity.cnmc.V21_2019_12_17.Cliente();

                    if (r["TipoIdentificador"] != DBNull.Value) cli.IdCliente.TipoIdentificador = r["TipoIdentificador"].ToString();
                    if (r["Identificador"] != DBNull.Value) cli.IdCliente.Identificador = r["Identificador"].ToString();
                    if (r["TipoPersona"] != DBNull.Value) cli.IdCliente.TipoPersona = r["TipoPersona"].ToString();

                    if (r["NombreDePila"] != DBNull.Value) cli.Nombre.NombreDePila = r["NombreDePila"].ToString();
                    if (r["PrimerApellido"] != DBNull.Value) cli.Nombre.PrimerApellido = r["PrimerApellido"].ToString();
                    if (r["SegundoApellido"] != DBNull.Value) cli.Nombre.SegundoApellido = r["SegundoApellido"].ToString();
                    if (r["RazonSocial"] != DBNull.Value) cli.Nombre.RazonSocial = r["RazonSocial"].ToString();

                    xml.CambiodeComercializadorSinCambios.Cliente = cli;


                }
                db.CloseConnection();
                return xml;
            }
            catch (Exception ex)
            {

                return null;
            }


        }
        public EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 CargaDatosM1(int id)  
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml =
                new EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101();
           ;

            try
            {

                strSql = "SELECT cab.CodigoREEEmpresaEmisora, cab.CodigoREEEmpresaDestino, cab.CodigoDelProceso,"
                    + " cab.CodigoDePaso, cab.CodigoDeSolicitud, cab.SecuencialDeSolicitud, cab.FechaSolicitud,"
                    + " cab.CUPS,sol.CNAE, sol.IndActivacion, sol.SolicitudTension, sol.TipoAutoconsumo, sol.TipoContratoATR,"
                    + " sol.TarifaATR,sol.Potencia1, sol.Potencia2, sol.Potencia3, sol.Potencia4, sol.Potencia5,"
                    + " sol.Potencia6, sol.ModoControlPotencia,con.PersonaDeContacto, con.PrefijoPais, con.Numero,"
                    + " cli.TipoIdentificador, cli.Identificador, cli.TipoPersona,cli.NombreDePila, cli.PrimerApellido,"
                    + " cli.SegundoApellido, cli.RazonSocial"
                    + " FROM cnmc_t_cabecera cab"
                    + " INNER JOIN cnmc_t_datossolicitudnew sol ON"
                    + " sol.id = cab.id"
                    + " INNER JOIN cnmc_t_cliente cli ON"
                    + " cli.id = cab.id"
                    + " LEFT JOIN cnmc_t_contacto con ON"
                    + " con.id = cab.id"
                    + " WHERE cab.id = " + id;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    xml.Cabecera.CodigoREEEmpresaEmisora = r["CodigoREEEmpresaEmisora"].ToString();
                    xml.Cabecera.CodigoREEEmpresaDestino = r["CodigoREEEmpresaDestino"].ToString();
                    xml.Cabecera.CodigoDelProceso = r["CodigoDelProceso"].ToString();
                    xml.Cabecera.CodigoDePaso = r["CodigoDePaso"].ToString();
                    xml.Cabecera.CodigoDeSolicitud = r["CodigoDeSolicitud"].ToString();
                    xml.Cabecera.SecuencialDeSolicitud = r["SecuencialDeSolicitud"].ToString();
                    xml.Cabecera.FechaSolicitud = r["FechaSolicitud"].ToString();
                    xml.Cabecera.CUPS = r["CUPS"].ToString();

                    // DatosSolicitud_C1
                   // var datos = xml.CambiodeComercializadorSinCambios.DatosSolicitud;
                   // if (r["CNAE"] != DBNull.Value) datos.cnae = r["CNAE"].ToString();
                   // if (r["IndActivacion"] != DBNull.Value) datos.indActivacion = r["IndActivacion"].ToString();
                   // if (r["SolicitudTension"] != DBNull.Value) datos.solicitudTension = r["SolicitudTension"].ToString();
                    // if (r["TipoModificacion"] != DBNull.Value) datos.tipoModificacion = r["TipoModificacion"].ToString();
                    //if (r["TipoSolicitudAdministrativa"] != DBNull.Value) datos.tipoSolicitudAdministrativa = r["TipoSolicitudAdministrativa"].ToString();
                    // if (r["FechaPrevistaAccion"] != DBNull.Value) datos.fechaPrevistaAccion = r["FechaPrevistaAccion"].ToString();
                    //  if (r["ContratacionIncondicionalPS"] != DBNull.Value) datos.contratacionIncondicionalPS = r["ContratacionIncondicionalPS"].ToString();
                    // if (r["ContratacionIncondicionalBS"] != DBNull.Value) datos.contratacionIncondicionalBS = r["ContratacionIncondicionalBS"].ToString();
                    // if (r["TensionSolicitada"] != DBNull.Value) datos.TensionSolicitada = r["TensionSolicitada"].ToString();

                    // Cliente
                    var cli = new EndesaEntity.cnmc.V21_2019_12_17.Cliente();

                    if (r["TipoIdentificador"] != DBNull.Value) cli.IdCliente.TipoIdentificador = r["TipoIdentificador"].ToString();
                    if (r["Identificador"] != DBNull.Value) cli.IdCliente.Identificador = r["Identificador"].ToString();
                    if (r["TipoPersona"] != DBNull.Value) cli.IdCliente.TipoPersona = r["TipoPersona"].ToString();

                    if (r["NombreDePila"] != DBNull.Value) cli.Nombre.NombreDePila = r["NombreDePila"].ToString();
                    if (r["PrimerApellido"] != DBNull.Value) cli.Nombre.PrimerApellido = r["PrimerApellido"].ToString();
                    if (r["SegundoApellido"] != DBNull.Value) cli.Nombre.SegundoApellido = r["SegundoApellido"].ToString();
                    if (r["RazonSocial"] != DBNull.Value) cli.Nombre.RazonSocial = r["RazonSocial"].ToString();

                   // xml.CambiodeComercializadorSinCambios.Cliente = cli;


                }
                db.CloseConnection();
                return xml;
            }
            catch (Exception ex)
            {

                return null;
            }


        }


        //---A301

        public void CreaMensajeA302_A(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA302_A a302a)
        {

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_A302_"
                + (a302a.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");


            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA302_A));
            serializer.Serialize(writer, a302a, ns);
            writer.Close();


            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_a302a"));
            if(mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void CreaMensajeA302_R(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA302_R a302r)
        {
           
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_A302_"
                + (a302r.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");

          
            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA302_R));
            serializer.Serialize(writer, a302r, ns);
            writer.Close();
           

            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_a302r"));
            if (mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void CreaMensajeA305(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 a301)
        {
            List<Potencia> lista_potencias = new List<Potencia>();

            EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305 xml = new TipoMensajeA305();
            xml.Cabecera.CodigoREEEmpresaEmisora = a301.Cabecera.CodigoREEEmpresaDestino;
            xml.Cabecera.CodigoREEEmpresaDestino = a301.Cabecera.CodigoREEEmpresaEmisora;
            xml.Cabecera.CodigoDelProceso = a301.Cabecera.CodigoDelProceso;
            xml.Cabecera.CodigoDePaso = "02";
            xml.Cabecera.CodigoDeSolicitud = a301.Cabecera.CodigoDeSolicitud;
            xml.Cabecera.SecuencialDeSolicitud = a301.Cabecera.SecuencialDeSolicitud;
            xml.Cabecera.FechaSolicitud = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss").ToString();
            xml.Cabecera.CUPS = a301.Cabecera.CUPS;
        }
        public void CreaMensajeA305v2(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305 a305)
        {

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_A305_"
                + (a305.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");


            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA305));
            serializer.Serialize(writer, a305, ns);
            writer.Close();


            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_a305"));
            if (mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //delete
            }

        }
        //-----C101_

        public void CreaMensajeC102_A(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_A c102a)
        {

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_C102_A"
                + (c102a.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");


            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_A));
            serializer.Serialize(writer, c102a, ns);
            writer.Close();


            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_c102a"));
            if (mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void CreaMensajeC102_R(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_R c102r)
        {

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_C102_r"
                + (c102r.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");


            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC102_R));
            serializer.Serialize(writer, c102r, ns);
            writer.Close();


            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_a302r"));
            if (mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void CreaMensajeC105(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC105 c105)
        {

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_C105_"
                + (c105.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");


            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC105));
            serializer.Serialize(writer, c105, ns);
            writer.Close();


            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_c105"));
            if (mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //delete
            }

        }

        //-------- M101
        public void CreaMensajeM102_A(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM102A m102a)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_M102_A"
                + (m102a.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");

            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM102A));
          
            serializer.Serialize(writer, m102a, ns);
            writer.Close();


            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_m102a"));
            if (mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void CreaMensajeM102_R(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM102_R m102r)
        {

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = Encoding.UTF8;
            string fichero = @"c:\Temp\XML_M102_R"
                + (m102r.Cabecera.CodigoDeSolicitud) + "_"
                + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";

            XmlWriter writer = XmlWriter.Create(fichero, settings);

            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "http://localhost/elegibilidad");


            XmlSerializer serializer = new XmlSerializer(typeof(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM102_R));
            serializer.Serialize(writer, m102r, ns);
            writer.Close();


            string mensaje = ValidateSchema(fichero, System.Environment.CurrentDirectory + p.GetValue("xsd_a302r"));
            if (mensaje == "")
            {
                MessageBox.Show("El XML sea ha generado correctamente en " + fichero.ToString(), "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El XML no es válido según el esquema: " + mensaje, "Genera XML", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        //
        public string ValidateSchema(string xmlPath, string xsdPath)
        {
           // string mensaje = "";
             string mensaje = string.Empty;
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlPath);

            xml.Schemas.Add(null, xsdPath);

            //try
            //{
            //    xml.Validate(null);
            //}
            //catch (XmlSchemaValidationException e)
              xml.Validate((sender, args) =>

              {
                  //  return e.Message;
                  //  }
                  //  return mensaje;
                  mensaje += args.Severity + ": " + args.Message + Environment.NewLine;
              });

            return mensaje.Trim();
        }


        public void Guarda_XML(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int id = 0;
            try
            {

                if (firstOnly)
                {
                    sb.Append("REPLACE INTO cnmc_t_cabecera");
                    sb.Append(" (id, CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,");
                    sb.Append(" CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud,");
                    sb.Append(" CUPS, created_by, created_date) values ");
                    firstOnly = false;
                }

                i++;

                id = GetID(xml.Cabecera.CodigoDelProceso, xml.Cabecera.CodigoDePaso,
                    xml.Cabecera.CodigoDeSolicitud);

                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaEmisora).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaDestino).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDelProceso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDePaso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDeSolicitud).Append("',");
                sb.Append("'").Append(xml.Cabecera.SecuencialDeSolicitud).Append("',");
                
                sb.Append("'").Append(Convert.ToDateTime(xml.Cabecera.FechaSolicitud).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.Cabecera.CUPS).Append("',");
                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");


                //irh
                Guarda_DatosSolicitud(id,xml);
                // Guarda_DatosSolicitudVa301(id,xml);

                Guarda_Contacto(id,xml);
                Guarda_Cliente(id,xml);


                if (i == 1)
                {
                   
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }

            }
            catch(Exception ex)
            {

            }
        }

        public void Guarda_XMLV30_a301(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int id = 0;
            try
            {

                if (firstOnly)
                {
                    sb.Append("REPLACE INTO cnmc_t_cabecera");
                    sb.Append(" (id, CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,");
                    sb.Append(" CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud,");
                    sb.Append(" CUPS, created_by, created_date) values ");
                    firstOnly = false;
                }

                i++;

                id = GetID(xml.Cabecera.CodigoDelProceso, xml.Cabecera.CodigoDePaso,
                    xml.Cabecera.CodigoDeSolicitud);

                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaEmisora).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaDestino).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDelProceso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDePaso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDeSolicitud).Append("',");
                sb.Append("'").Append(xml.Cabecera.SecuencialDeSolicitud).Append("',");

                sb.Append("'").Append(Convert.ToDateTime(xml.Cabecera.FechaSolicitud).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.Cabecera.CUPS).Append("',");
                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
              //sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("')");

                //irh
                // Guarda_DatosSolicitud(id, xml);
                Guarda_DatosSolicitudV30(id,xml);

               //Guarda_Contacto(id, xml);
                Guarda_ContactoV30(id, xml);
               // Guarda_Cliente(id, xml);
                Guarda_ClienteV30(id, xml);

                if (i == 1)
                {

                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString(), db.con);
                //  command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }

            }
            catch (Exception ex)
            {

            }
        }
        public void Guarda_XML(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA302_A xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int id = 0;
            try
            {

                if (firstOnly)
                {
                    sb.Append("REPLACE INTO cnmc_t_cabecera");
                    sb.Append(" (id, CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,");
                    sb.Append(" CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud,");
                    sb.Append(" CUPS, created_by, created_date) values ");
                    firstOnly = false;
                }

                i++;

                id = GetID(xml.Cabecera.CodigoDelProceso, xml.Cabecera.CodigoDePaso,
                    xml.Cabecera.CodigoDeSolicitud);

                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaEmisora).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaDestino).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDelProceso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDePaso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDeSolicitud).Append("',");
                sb.Append("'").Append(xml.Cabecera.SecuencialDeSolicitud).Append("',");

                sb.Append("'").Append(Convert.ToDateTime(xml.Cabecera.FechaSolicitud).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.Cabecera.CUPS).Append("',");
                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                //Guarda_DatosSolicitud(id, xml);
                
                if (i == 1)
                {

                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }

            }
            catch (Exception ex)
            {

            }
        }

        public void Guarda_XMLc101(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeC101 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int id = 0;
            try
            {
                if (firstOnly)
                {
                    sb.Append("REPLACE INTO cnmc_t_cabecera");
                    sb.Append(" (id, CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,");
                    sb.Append(" CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud,");
                    sb.Append(" CUPS, created_by, created_date) values ");
                    firstOnly = false;
                }
                i++;
                id = GetID(xml.Cabecera.CodigoDelProceso, xml.Cabecera.CodigoDePaso,
                    xml.Cabecera.CodigoDeSolicitud);
                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaEmisora).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaDestino).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDelProceso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDePaso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDeSolicitud).Append("',");
                sb.Append("'").Append(xml.Cabecera.SecuencialDeSolicitud).Append("',");
                sb.Append("'").Append(Convert.ToDateTime(xml.Cabecera.FechaSolicitud).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.Cabecera.CUPS).Append("',");
                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

            }
            catch (Exception ex)
            {
                // Manejo de excepciones
            }
            finally
            {
                if (i == 1)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }
            }
        }

        public void Guarda_XMLb101(EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeB101 xml) // ojo version 30 y v21
        {
        }

        public void Guarda_XMLc201(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeC201 xml) 
        {
        }

        public void Guarda_XMLm101REVIASAR(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml)
        {

            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int i = 0;
            int id = 0;
            try
            {

                if (firstOnly)
                {
                    sb.Append("REPLACE INTO cnmc_t_cabecera");
                    sb.Append(" (id, CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso,");
                    sb.Append(" CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud,");
                    sb.Append(" CUPS, created_by, created_date) values ");
                    firstOnly = false;
                }

                i++;

                id = GetID(xml.Cabecera.CodigoDelProceso, xml.Cabecera.CodigoDePaso,
                    xml.Cabecera.CodigoDeSolicitud);

                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaEmisora).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoREEEmpresaDestino).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDelProceso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDePaso).Append("',");
                sb.Append("'").Append(xml.Cabecera.CodigoDeSolicitud).Append("',");
                sb.Append("'").Append(xml.Cabecera.SecuencialDeSolicitud).Append("',");

                sb.Append("'").Append(Convert.ToDateTime(xml.Cabecera.FechaSolicitud).ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                sb.Append("'").Append(xml.Cabecera.CUPS).Append("',");
                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                //sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("')");

                //irh
                 Guarda_DatosSolicitudM1(id, xml);
                
                 Guarda_ContactoM1(id, xml);
                
                 Guarda_ClienteM1(id, xml);

                if (i == 1)
                {    

                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString(), db.con);
                    //  command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    i = 0;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar cliente  Guarda_XMLm101: " + ex.Message);
            }
        }

        public void Guarda_XMLm101(EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            int id = 0;

            try
            {
                // Generar sentencia para cnmc_t_cabecera
                sb.Append("REPLACE INTO cnmc_t_cabecera ");
                sb.Append("(id, CodigoREEEmpresaEmisora, CodigoREEEmpresaDestino, CodigoDelProceso, ");
                sb.Append("CodigoDePaso, CodigoDeSolicitud, SecuencialDeSolicitud, FechaSolicitud, ");
                sb.Append("CUPS, created_by, created_date) VALUES (");

                id = GetID(xml.Cabecera.CodigoDelProceso, xml.Cabecera.CodigoDePaso, xml.Cabecera.CodigoDeSolicitud);

                sb.Append(id).Append(",");
                sb.Append($"'{xml.Cabecera.CodigoREEEmpresaEmisora}',");
                sb.Append($"'{xml.Cabecera.CodigoREEEmpresaDestino}',");
                sb.Append($"'{xml.Cabecera.CodigoDelProceso}',");
                sb.Append($"'{xml.Cabecera.CodigoDePaso}',");
                sb.Append($"'{xml.Cabecera.CodigoDeSolicitud}',");
                sb.Append($"'{xml.Cabecera.SecuencialDeSolicitud}',");
                sb.Append($"'{Convert.ToDateTime(xml.Cabecera.FechaSolicitud):yyyy-MM-dd HH:mm:ss}',");
                sb.Append($"'{xml.Cabecera.CUPS}',");
                sb.Append($"'{Environment.UserName.ToUpper()}',");
                sb.Append($"'{DateTime.Now:yyyy-MM-dd HH:mm:ss}'");
                sb.Append(")");

                // Ejecutar inserción en cnmc_t_cabecera
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                // Guardar datos auxiliares
                Guarda_DatosSolicitudM1(id, xml);
                Guarda_ContactoM1(id, xml);
                Guarda_ClienteM1(id, xml);
            }
            catch (MySqlException dbEx)
            {
                MessageBox.Show("Error de base de datos al guardar M101: " + dbEx.Message, "MySQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error general al guardar M101: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        //--- solicitudes  --

        private void Guarda_DatosSolicitud(int id, EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int num_potencias = 0;

            sb.Append("REPLACE INTO cnmc_t_datossolicitud");
            sb.Append(" (id, CNAE, IndActivacion, SolicitudTension, TipoAutoconsumo,");
            sb.Append(" TipoContratoATR, TarifaATR, Potencia1, Potencia2, Potencia3,");
            sb.Append(" Potencia4, Potencia5, Potencia6, ModoControlPotencia,");
            sb.Append(" created_by, created_date) values ");

            sb.Append("(").Append(id).Append(",");

            if (xml.Alta.DatosSolicitud.cnae != null)
                sb.Append("'").Append(xml.Alta.DatosSolicitud.cnae).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.DatosSolicitud.IndActivacion != null)
                sb.Append("'").Append(xml.Alta.DatosSolicitud.IndActivacion).Append("',");
            else
                sb.Append("null,");

            sb.Append("null,"); // SolicitudTension

            if (xml.Alta.Contrato.TipoAutoconsumo != null)
                sb.Append("'").Append(xml.Alta.Contrato.TipoAutoconsumo).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Contrato.TipoContratoATR != null)
                sb.Append("'").Append(xml.Alta.Contrato.TipoContratoATR).Append("',");
            else
                sb.Append("null,");


            if (xml.Alta.Contrato.CondicionesContractuales.TarifaATR != null)
                sb.Append("'").Append(xml.Alta.Contrato.CondicionesContractuales.TarifaATR).Append("',");
            else
                sb.Append("null,");


            foreach (EndesaEntity.cnmc.V21_2019_12_17.Potencia p in
                xml.Alta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia)
            {
                
                if (p.periodo != null)
                    sb.Append(p.potencia).Append(",");
                else
                    sb.Append("null,");
                
            }

            for(int i = xml.Alta.Contrato.CondicionesContractuales.PotenciasContratadas.Potencia.Count; i < 6; i++)
                sb.Append("null,");

            if (xml.Alta.Contrato.CondicionesContractuales.ModoControlPotencia != null)
                sb.Append("'").Append(xml.Alta.Contrato.CondicionesContractuales.ModoControlPotencia).Append("',");
            else
                sb.Append("null,");

            sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
            sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();


        }


        //        private void Guarda_DatosSolicitudV30(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml)  //  irh
        //        //  private void Guarda_DatosSolicitudVa301(int id, EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml)
        //        {
        //            servidores.MySQLDB db;
        //            MySqlCommand command;
        //            StringBuilder sb = new StringBuilder();

        //            sb.Append("REPLACE INTO cnmc_t_datossolicitudnew (");
        //            sb.Append("id, CNAE, IndActivacion, SolicitudTension, TipoAutoconsumo, TipoContratoATR, TarifaATR, ");
        //            sb.Append("IndEsencial,fechaPrevistaAccion,TensionSolicitada, ");
        //            sb.Append("Potencia1, Potencia2, Potencia3, Potencia4, Potencia5, Potencia6, ModoControlPotencia, ");
        //            sb.Append("CodContrato, Fecha_finalizacion, TipoActivacionPrevista, TipoCUPS, cau, TipoSubseccion, ");
        //            sb.Append("Colectivo, PotInstaladaGen, TipoInstalacion, SSAA, UnicoContrato, RefCatastro , ");
        //            sb.Append("CIL, TecGenerado, CUPSPrincipal,FechaActivacionPrevista ,TensionDelSuministro, ");
        //            sb.Append("VAsTrafo, PeriodicidadFacturacion, TipoTelegestion,  ");
        //            sb.Append("created_by, created_date) VALUES (");

        //            var datos = xml.Alta.DatosSolicitud; //  ojo version 30 _  AltaA301 Alta
        //            var contrato = xml.Alta.Contrato;
        //            var cond = xml.Alta.Contrato.CondicionesContractuales;  // contrato.CondicionesContractuales;
        //            AutoconsumoSolicitudAlta autoconsumo = null;
        //            DatosCAUAlta cau = null;
        //            DatosInstGenSolicitud DatosInstGen = null;

        //            if (xml.Alta.Contrato.Autoconsumo != null)
        //            {
        //                autoconsumo = xml.Alta.Contrato.Autoconsumo;
        //                sb.Append(autoconsumo.DatosSuministro.TipoCUPS != null ? $"'{autoconsumo.DatosSuministro.TipoCUPS}'," : "null,");

        //                sb.Append(autoconsumo.DatosCAU.TipoSubseccion != null ? $"'{autoconsumo.DatosCAU.TipoSubseccion}'," : "null,");

        //                cau = xml.Alta.Contrato.Autoconsumo.DatosCAU;
        //                sb.Append(cau.TipoAutoconsumo != null ? $"'{cau.TipoAutoconsumo}'," : "null,");
        //                sb.Append(cau.Colectivo != null ? $"'{cau.Colectivo}'," : "null,");
        //                sb.Append(autoconsumo.DatosCAU.CAU != null ? $"'{autoconsumo.DatosCAU.CAU}'," : "null,");

        //                DatosInstGen = xml.Alta.Contrato.Autoconsumo.DatosCAU.DatosInstGen;
        //                sb.Append(DatosInstGen.PotInstaladaGen != null ? $"{DatosInstGen.PotInstaladaGen}," : "null,");
        //                sb.Append(DatosInstGen.TipoInstalacion != null ? $"'{DatosInstGen.TipoInstalacion}'," : "null,");
        //                sb.Append(DatosInstGen.SSAA != null ? $"'{DatosInstGen.SSAA}'," : "null,");
        //                sb.Append(DatosInstGen.UnicoContrato != null ? $"'{DatosInstGen.UnicoContrato}'," : "null,");
        //                sb.Append(DatosInstGen.RefCatastro != null ? $"'{DatosInstGen.RefCatastro}'," : "null,");

        //                sb.Append(DatosInstGen.CIL != null ? $"'{DatosInstGen.CIL}'," : "null,");
        //                sb.Append(DatosInstGen.TecGenerador != null ? $"'{DatosInstGen.TecGenerador}'," : "null,");


        //            }
        //            sb.Append(id).Append(",");
        //            sb.Append(datos.cnae != null ? $"'{datos.cnae}'," : "null,");
        //            sb.Append(datos.IndActivacion != null ? $"'{datos.IndActivacion}'," : "null,");
        //            sb.Append(datos.SolicitudTension != null ? $"'{datos.SolicitudTension}'," : "null,");

        //            sb.Append(contrato.TipoContratoATR != null ? $"'{contrato.TipoContratoATR}'," : "null,");
        //            sb.Append(cond.TarifaATR != null ? $"'{cond.TarifaATR}'," : "null,");   
        //            //
        //            sb.Append(datos.IndEsencial != null ? $"'{datos.IndEsencial}'," : "null,");
        //            sb.Append(datos.fechaPrevistaAccion != null ? $"'{datos.fechaPrevistaAccion}'," : "null,");
        //            sb.Append(datos.SolicitudTension != null ? $"'{datos.SolicitudTension}'," : "null,");
        //            sb.Append(datos.TensionSolicitada != null ? $"'{datos.TensionSolicitada}'," : "null,");
        //            //
        //            foreach (var p in cond.PotenciasContratadas.Potencia)
        //                sb.Append(p.periodo != null ? $"{p.potencia}," : "null,");
        //            for (int i = cond.PotenciasContratadas.Potencia.Count; i < 6; i++)
        //                sb.Append("null,");
        //            sb.Append(cond.ModoControlPotencia != null ? $"'{cond.ModoControlPotencia}'," : "null,");
        //            //-- null
        //            // sb.Append(contrato.IdContrato.CodContrato != null ? $"'{contrato.IdContrato.CodContrato}'," : "null,");
        //            sb.Append(
        //                     contrato.IdContrato != null && contrato.IdContrato.CodContrato != null
        //                     ? $"'{contrato.IdContrato.CodContrato}',"
        //                     : "null,"
        //);
        //            //
        //            sb.Append(contrato.FechaFinalizacion != null ? $"'{contrato.FechaFinalizacion}'," : "null,");
        //            sb.Append(contrato.TipoActivacionPrevista != null ? $"'{contrato.TipoActivacionPrevista}'," : "null,");

        //            // sb.Append("Colectivo, PotInstaladaGen, TipoInstalacion, SSAA, UnicoContrato, MotivoTraspaso, ");

        //            sb.Append(contrato.CUPSPrincipal != null ? $"'{contrato.CUPSPrincipal}'," : "null,");
        //            sb.Append(contrato.FechaActivacionPrevista != null ? $"'{contrato.FechaActivacionPrevista}'," : "null,");
        //            sb.Append(cond.TensionDelSuministro != null ? $"'{cond.TensionDelSuministro}'," : "null,");

        //            sb.Append(cond.VAsTrafo != null ? $"'{cond.VAsTrafo}'," : "null,");
        //            sb.Append(cond.PeriodicidadFacturacion != null ? $"'{cond.PeriodicidadFacturacion}'," : "null,");
        //            sb.Append(cond.TipoTelegestion != null ? $"'{cond.TipoTelegestion}'," : "null,");
        //            sb.Append(datos.TensionSolicitada != null ? $"'{datos.TensionSolicitada}'," : "null,");

        //            sb.Append($"'{Environment.UserName.ToUpper()}',");
        //            sb.Append($"'{DateTime.Now:yyyy-MM-dd HH:mm:ss}'");

        //            sb.Append(")");

        //            db = new MySQLDB(MySQLDB.Esquemas.CON);
        //            command = new MySqlCommand(sb.ToString(), db.con);
        //            command.ExecuteNonQuery();
        //            db.CloseConnection();

        //        }
        private void Guarda_DatosSolicitudV30(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();

            sb.Append("REPLACE INTO cnmc_t_datossolicitudnew (");
            sb.Append("id, CNAE, IndActivacion, SolicitudTension, ");
            sb.Append("TipoAutoconsumo, TipoCUPS, cau, TipoSubseccion, Colectivo, ");
            sb.Append("PotInstaladaGen, TipoInstalacion, SSAA, UnicoContrato, RefCatastro, ");
            sb.Append("CIL, TecGenerado, ");
            sb.Append("TipoContratoATR, TarifaATR, IndEsencial, fechaPrevistaAccion, TensionSolicitada, ");
            sb.Append("Potencia1, Potencia2, Potencia3, Potencia4, Potencia5, Potencia6, ");
            sb.Append("ModoControlPotencia, CodContrato, Fecha_finalizacion, TipoActivacionPrevista, ");
            sb.Append("CUPSPrincipal, FechaActivacionPrevista, TensionDelSuministro, ");
            sb.Append("VAsTrafo, PeriodicidadFacturacion, TipoTelegestion, ");
            sb.Append("created_by, created_date) VALUES (");

            var datos = xml.Alta.DatosSolicitud;
            var contrato = xml.Alta.Contrato;
            var cond = contrato.CondicionesContractuales;

            // id
            sb.Append($"{id},");
            // CNAE
            sb.Append(datos?.cnae != null ? $"'{datos.cnae}'," : "null,");
            // IndActivacion
            sb.Append(datos?.IndActivacion != null ? $"'{datos.IndActivacion}'," : "null,");
            // SolicitudTension
            sb.Append(datos?.SolicitudTension != null ? $"'{datos.SolicitudTension}'," : "null,");

            // ---------- AUTOCONSUMO ----------
            if (contrato.Autoconsumo != null)
            {
                var auto = contrato.Autoconsumo;

                sb.Append(auto.DatosCAU?.TipoAutoconsumo != null
                    ? $"'{auto.DatosCAU.TipoAutoconsumo}'," : "null,");

                sb.Append(auto.DatosSuministro?.TipoCUPS != null
                    ? $"'{auto.DatosSuministro.TipoCUPS}'," : "null,");

                sb.Append(auto.DatosCAU?.CAU != null
                    ? $"'{auto.DatosCAU.CAU}'," : "null,");

                sb.Append(auto.DatosCAU?.TipoSubseccion != null
                    ? $"'{auto.DatosCAU.TipoSubseccion}'," : "null,");

                sb.Append(auto.DatosCAU?.Colectivo != null
                    ? $"'{auto.DatosCAU.Colectivo}'," : "null,");

                var datosInstGen = auto.DatosCAU?.DatosInstGen;

                sb.Append(datosInstGen?.PotInstaladaGen != null
                    ? $"{datosInstGen.PotInstaladaGen}," : "null,");

                sb.Append(datosInstGen?.TipoInstalacion != null
                    ? $"'{datosInstGen.TipoInstalacion}'," : "null,");

                sb.Append(datosInstGen?.SSAA != null
                    ? $"'{datosInstGen.SSAA}'," : "null,");

                sb.Append(datosInstGen?.UnicoContrato != null
                    ? $"'{datosInstGen.UnicoContrato}'," : "null,");

                sb.Append(datosInstGen?.RefCatastro != null
                    ? $"'{datosInstGen.RefCatastro}'," : "null,");

                sb.Append(datosInstGen?.CIL != null
                    ? $"'{datosInstGen.CIL}'," : "null,");

                sb.Append(datosInstGen?.TecGenerador != null
                    ? $"'{datosInstGen.TecGenerador}'," : "null,");
            }
            else
            {
                // Si Autoconsumo es null, meter NULLs para todos esos campos
                sb.Append("null,null,null,null,null,");
                sb.Append("null,null,null,null,null,");
                sb.Append("null,null,");
            }

            // TipoContratoATR
            sb.Append(contrato?.TipoContratoATR != null ? $"'{contrato.TipoContratoATR}'," : "null,");
            // TarifaATR
            sb.Append(cond?.TarifaATR != null ? $"'{cond.TarifaATR}'," : "null,");
            // IndEsencial
            sb.Append(datos?.IndEsencial != null ? $"'{datos.IndEsencial}'," : "null,");
            // fechaPrevistaAccion
            sb.Append(datos?.fechaPrevistaAccion != null ? $"'{datos.fechaPrevistaAccion}'," : "null,");
            // TensionSolicitada
            sb.Append(datos?.TensionSolicitada != null ? $"'{datos.TensionSolicitada}'," : "null,");

            // Potencias
            if (cond?.PotenciasContratadas?.Potencia != null)
            {
                foreach (var p in cond.PotenciasContratadas.Potencia)
                    sb.Append(p?.potencia != null ? $"{p.potencia}," : "null,");

                for (int i = cond.PotenciasContratadas.Potencia.Count; i < 6; i++)
                    sb.Append("null,");
            }
            else
            {
                // Si no hay ninguna potencia, 6 nulls
                for (int i = 0; i < 6; i++)
                    sb.Append("null,");
            }

            // ModoControlPotencia
            sb.Append(cond?.ModoControlPotencia != null ? $"'{cond.ModoControlPotencia}'," : "null,");

            // CodContrato
            sb.Append(contrato?.IdContrato?.CodContrato != null
                ? $"'{contrato.IdContrato.CodContrato}'," : "null,");
            // Fecha_finalizacion
            sb.Append(contrato?.FechaFinalizacion != null ? $"'{contrato.FechaFinalizacion}'," : "null,");
            // TipoActivacionPrevista
            sb.Append(contrato?.TipoActivacionPrevista != null ? $"'{contrato.TipoActivacionPrevista}'," : "null,");
            // CUPSPrincipal
            sb.Append(contrato?.CUPSPrincipal != null ? $"'{contrato.CUPSPrincipal}'," : "null,");
            // FechaActivacionPrevista
            sb.Append(contrato?.FechaActivacionPrevista != null ? $"'{contrato.FechaActivacionPrevista}'," : "null,");
            // TensionDelSuministro
            sb.Append(cond?.TensionDelSuministro != null ? $"'{cond.TensionDelSuministro}'," : "null,");
            // VAsTrafo
            sb.Append(cond?.VAsTrafo != null ? $"'{cond.VAsTrafo}'," : "null,");
            // PeriodicidadFacturacion
            sb.Append(cond?.PeriodicidadFacturacion != null ? $"'{cond.PeriodicidadFacturacion}'," : "null,");
            // TipoTelegestion
            sb.Append(cond?.TipoTelegestion != null ? $"'{cond.TipoTelegestion}'," : "null,");

            // created_by
            sb.Append($"'{Environment.UserName.ToUpper()}',");
            // created_date
            sb.Append($"'{DateTime.Now:yyyy-MM-dd HH:mm:ss}'");

            sb.Append(")");

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(sb.ToString(), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
        //m101
        //private void Guarda_DatosSolicitudM1(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml)
        //{
        //    var datos = xml.ModificacionDeATR.DatosSolicitud;
        //    var contrato = xml.ModificacionDeATR.Contrato;
        //    var cond = contrato.CondicionesContractuales;

        //    StringBuilder sb = new StringBuilder();
        //    sb.Append("REPLACE INTO cnmc_t_datossolicitudnew (");
        //    sb.Append("id, CNAE, IndActivacion, SolicitudTension, ");
        //    sb.Append("TipoAutoconsumo, TipoCUPS, cau, TipoSubseccion, Colectivo, ");
        //    sb.Append("PotInstaladaGen, TipoInstalacion, SSAA, UnicoContrato, RefCatastro, ");
        //    sb.Append("CIL, TecGenerado, ");
        //    sb.Append("TipoContratoATR, TarifaATR, IndEsencial, fechaPrevistaAccion, TensionSolicitada, ");
        //    sb.Append("Potencia1, Potencia2, Potencia3, Potencia4, Potencia5, Potencia6, ");
        //    sb.Append("ModoControlPotencia, CodContrato, Fecha_finalizacion, TipoActivacionPrevista, ");
        //    sb.Append("CUPSPrincipal, FechaActivacionPrevista, TensionDelSuministro, ");
        //    sb.Append("VAsTrafo, PeriodicidadFacturacion, TipoTelegestion, ");
        //    sb.Append("created_by, created_date) VALUES (");

        //    sb.Append($"{id},");
        //   // sb.Append(datos?.cnae != null ? $"'{datos.cnae}'," : "null,");
        //    sb.Append(datos?.IndActivacion != null ? $"'{datos.IndActivacion}'," : "null,");
        //   // sb.Append(datos?.SolicitudTension != null ? $"'{datos.SolicitudTension}'," : "null,");

        //    if (contrato.Autoconsumo != null)
        //    {
        //        var auto = contrato.Autoconsumo;

        //        sb.Append(auto.DatosCAU?.TipoAutoconsumo != null ? $"'{auto.DatosCAU.TipoAutoconsumo}'," : "null,");
        //        sb.Append(auto.DatosSuministro?.TipoCUPS != null ? $"'{auto.DatosSuministro.TipoCUPS}'," : "null,");
        //        sb.Append(auto.DatosCAU?.CAU != null ? $"'{auto.DatosCAU.CAU}'," : "null,");
        //        sb.Append(auto.DatosCAU?.TipoSubseccion != null ? $"'{auto.DatosCAU.TipoSubseccion}'," : "null,");
        //        sb.Append(auto.DatosCAU?.Colectivo != null ? $"'{auto.DatosCAU.Colectivo}'," : "null,");

        //        var datosInstGen = auto.DatosCAU?.DatosInstGen;

        //        sb.Append(datosInstGen?.PotInstaladaGen != null ? $"{datosInstGen.PotInstaladaGen}," : "null,");
        //        sb.Append(datosInstGen?.TipoInstalacion != null ? $"'{datosInstGen.TipoInstalacion}'," : "null,");
        //        sb.Append(datosInstGen?.SSAA != null ? $"'{datosInstGen.SSAA}'," : "null,");
        //        sb.Append(datosInstGen?.UnicoContrato != null ? $"'{datosInstGen.UnicoContrato}'," : "null,");
        //        sb.Append(datosInstGen?.RefCatastro != null ? $"'{datosInstGen.RefCatastro}'," : "null,");
        //        sb.Append(datosInstGen?.CIL != null ? $"'{datosInstGen.CIL}'," : "null,");
        //        sb.Append(datosInstGen?.TecGenerador != null ? $"'{datosInstGen.TecGenerador}'," : "null,");
        //    }
        //    else
        //    {
        //        sb.Append("null,null,null,null,null,null,null,null,null,null,null,null,null,");
        //    }

        //    sb.Append(contrato?.TipoContratoATR != null ? $"'{contrato.TipoContratoATR}'," : "null,");
        //    sb.Append(cond?.TarifaATR != null ? $"'{cond.TarifaATR}'," : "null,");
        //    sb.Append(datos?.IndEsencial != null ? $"'{datos.IndEsencial}'," : "null,");
        //   // sb.Append(datos?.fechaPrevistaAccion != null ? $"'{datos.fechaPrevistaAccion}'," : "null,");
        //   // sb.Append(datos?.TensionSolicitada != null ? $"'{datos.TensionSolicitada}'," : "null,");

        //    if (cond?.PotenciasContratadas?.Potencia != null)
        //    {
        //        foreach (var p in cond.PotenciasContratadas.Potencia)
        //            sb.Append(p?.potencia != null ? $"{p.potencia}," : "null,");

        //        for (int i = cond.PotenciasContratadas.Potencia.Count; i < 6; i++)
        //            sb.Append("null,");
        //    }
        //    else
        //    {
        //        for (int i = 0; i < 6; i++)
        //            sb.Append("null,");
        //    }

        //    sb.Append(cond?.ModoControlPotencia != null ? $"'{cond.ModoControlPotencia}'," : "null,");
        //   // sb.Append(contrato?.IdContrato?.CodContrato != null ? $"'{contrato.IdContrato.CodContrato}'," : "null,");
        //    sb.Append(contrato?.FechaFinalizacion != null ? $"'{contrato.FechaFinalizacion}'," : "null,");
        //   // sb.Append(contrato?.TipoActivacionPrevista != null ? $"'{contrato.TipoActivacionPrevista}'," : "null,");
        //    sb.Append(contrato?.CUPSPrincipal != null ? $"'{contrato.CUPSPrincipal}'," : "null,");
        //  //  sb.Append(contrato?.FechaActivacionPrevista != null ? $"'{contrato.FechaActivacionPrevista}'," : "null,");
        //    sb.Append(cond?.TensionDelSuministro != null ? $"'{cond.TensionDelSuministro}'," : "null,");
        //    sb.Append(cond?.VAsTrafo != null ? $"'{cond.VAsTrafo}'," : "null,");
        //    sb.Append(cond?.PeriodicidadFacturacion != null ? $"'{cond.PeriodicidadFacturacion}'," : "null,");
        //    sb.Append(cond?.TipoTelegestion != null ? $"'{cond.TipoTelegestion}'," : "null,");

        //    sb.Append($"'{Environment.UserName.ToUpper()}',");
        //    sb.Append($"'{DateTime.Now:yyyy-MM-dd HH:mm:ss}')");

        //    var db = new MySQLDB(MySQLDB.Esquemas.CON);
        //    var command = new MySqlCommand(sb.ToString(), db.con);
        //    command.ExecuteNonQuery();
        //    db.CloseConnection();
        //}
        private void Guarda_DatosSolicitudM1(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml)
        {
            try
            {
                var datos = xml.ModificacionDeATR.DatosSolicitud;
                var contrato = xml.ModificacionDeATR.Contrato;
                var cond = contrato.CondicionesContractuales;
                var auto = contrato.Autoconsumo;
                var cau = auto?.DatosCAU;
                var inst = cau?.DatosInstGen;

                var sb = new StringBuilder();
                sb.Append("REPLACE INTO cnmc_t_datossolicitudnew (");
                sb.Append("id, CNAE, IndActivacion, SolicitudTension, ");
                sb.Append("TipoAutoconsumo, TipoCUPS, cau, TipoSubseccion, Colectivo, ");
                sb.Append("PotInstaladaGen, TipoInstalacion, SSAA, UnicoContrato, RefCatastro, ");
                sb.Append("CIL, TecGenerado, ");
                sb.Append("TipoContratoATR, TarifaATR, IndEsencial, fechaPrevistaAccion, TensionSolicitada, ");
                sb.Append("Potencia1, Potencia2, Potencia3, Potencia4, Potencia5, Potencia6, ");
                sb.Append("ModoControlPotencia, CodContrato, Fecha_finalizacion, TipoActivacionPrevista, ");
                sb.Append("CUPSPrincipal, FechaActivacionPrevista, TensionDelSuministro, ");
                sb.Append("VAsTrafo, PeriodicidadFacturacion, TipoTelegestion, ");
                sb.Append("created_by, created_date) VALUES (");

                sb.Append(id).Append(",");

                // DatosSolicitud_M1
                sb.Append(datos?.CNAE != null ? $"'{datos.CNAE}'," : "null,");
                sb.Append(datos?.IndActivacion != null ? $"'{datos.IndActivacion}'," : "null,");
                sb.Append("null,"); // SolicitudTension no existe

                // Autoconsumo
                sb.Append(cau?.TipoAutoconsumo != null ? $"'{cau.TipoAutoconsumo}'," : "null,");
                sb.Append(auto?.DatosSuministro?.TipoCUPS != null ? $"'{auto.DatosSuministro.TipoCUPS}'," : "null,");
                sb.Append(cau?.CAU != null ? $"'{cau.CAU}'," : "null,");
                sb.Append(cau?.TipoSubseccion != null ? $"'{cau.TipoSubseccion}'," : "null,");
                sb.Append(cau?.Colectivo != null ? $"'{cau.Colectivo}'," : "null,");

                sb.Append(inst?.PotInstaladaGen != null ? $"{inst.PotInstaladaGen}," : "null,");
                sb.Append(inst?.TipoInstalacion != null ? $"'{inst.TipoInstalacion}'," : "null,");
                sb.Append(inst?.SSAA != null ? $"'{inst.SSAA}'," : "null,");
                sb.Append(inst?.UnicoContrato != null ? $"'{inst.UnicoContrato}'," : "null,");
                sb.Append(inst?.RefCatastro != null ? $"'{inst.RefCatastro}'," : "null,");
                sb.Append(inst?.CIL != null ? $"'{inst.CIL}'," : "null,");
                sb.Append(inst?.TecGenerador != null ? $"'{inst.TecGenerador}'," : "null,");

                // Contrato
                sb.Append(contrato?.TipoContratoATR != null ? $"'{contrato.TipoContratoATR}'," : "null,");
                sb.Append(cond?.TarifaATR != null ? $"'{cond.TarifaATR}'," : "null,");
                sb.Append(datos?.IndEsencial != null ? $"'{datos.IndEsencial}'," : "null,");
                sb.Append(datos?.FechaPrevistaAccion != null ? $"'{datos.FechaPrevistaAccion}'," : "null,");
                sb.Append("null,"); // TensionSolicitada no existe

                // Potencias contratadas
                if (cond?.PotenciasContratadas?.Potencia != null)
                {
                    foreach (var p in cond.PotenciasContratadas.Potencia)
                        sb.Append(p?.potencia != null ? $"{p.potencia}," : "null,");
                    for (int i = cond.PotenciasContratadas.Potencia.Count; i < 6; i++)
                        sb.Append("null,");
                }
                else
                {
                    for (int i = 0; i < 6; i++) sb.Append("null,");
                }

                sb.Append(cond?.ModoControlPotencia != null ? $"'{cond.ModoControlPotencia}'," : "null,");
                sb.Append("null,"); // CodContrato NO existe en Contrato_M1
                //sb.Append(contrato?.IdContrato?.CodContrato != null ? $"'{contrato.IdContrato.CodContrato}'," : "null,");
                sb.Append(contrato?.FechaFinalizacion != null ? $"'{contrato.FechaFinalizacion}'," : "null,");
                sb.Append("null,"); //C NO existe en Contrato_M1
               // sb.Append(contrato?.TipoActivacionPrevista != null ? $"'{contrato.TipoActivacionPrevista}'," : "null,");
                sb.Append(contrato?.CUPSPrincipal != null ? $"'{contrato.CUPSPrincipal}'," : "null,");
                sb.Append("null,"); // C NO existe en Contrato_M1
              //  sb.Append(contrato?.FechaActivacionPrevista != null ? $"'{contrato.FechaActivacionPrevista}'," : "null,");
                sb.Append(cond?.TensionDelSuministro != null ? $"'{cond.TensionDelSuministro}'," : "null,");
                sb.Append(cond?.VAsTrafo != null ? $"'{cond.VAsTrafo}'," : "null,");
                sb.Append(cond?.PeriodicidadFacturacion != null ? $"'{cond.PeriodicidadFacturacion}'," : "null,");
                sb.Append(cond?.TipoTelegestion != null ? $"'{cond.TipoTelegestion}'," : "null,");

                // metadatos
                sb.Append($"'{Environment.UserName.ToUpper()}',");
                sb.Append($"'{DateTime.Now:yyyy-MM-dd HH:mm:ss}'");
                sb.Append(")");

                // Validación previa: contar columnas y valores
                var sql = sb.ToString();
                int colCount = sql.Substring(sql.IndexOf('(') + 1, sql.IndexOf(")") - sql.IndexOf("(") - 1).Split(',').Length;
                int valCount = sql.Substring(sql.LastIndexOf('(') + 1, sql.LastIndexOf(")") - sql.LastIndexOf('(') - 1).Split(',').Length;

                if (colCount != valCount)
                {
                    MessageBox.Show($"ERROR DE ESTRUCTURA: columnas={colCount}, valores={valCount}\n\n{sql}", "Estructura SQL incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var db = new MySQLDB(MySQLDB.Esquemas.CON);
                var command = new MySqlCommand(sql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar datos de solicitud M1:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Guarda_ContactoM1(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml)
        {
            var contacto = xml.ModificacionDeATR.Contrato.Contacto;
            if (contacto?.PersonaDeContacto == null) return;

            var sb = new StringBuilder();
            sb.Append("REPLACE INTO cnmc_t_contacto ");
            sb.Append("(id, PersonaDeContacto, PrefijoPais, Numero, created_by, created_date) VALUES (");
            sb.Append(id).Append(",");
            sb.Append($"'{contacto.PersonaDeContacto}',");
            sb.Append(contacto.Telefono?.PrefijoPais != null ? $"'{contacto.Telefono.PrefijoPais}'," : "null,");
            sb.Append(contacto.Telefono?.Numero != null ? $"'{contacto.Telefono.Numero}'," : "null,");
            sb.Append($"'{Environment.UserName.ToUpper()}',");
            sb.Append($"'{DateTime.Now:yyyy-MM-dd HH:mm:ss}')");

            var db = new MySQLDB(MySQLDB.Esquemas.CON);
            var command = new MySqlCommand(sb.ToString(), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void Guarda_ClienteM1(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeM101 xml)
        {
            var cli = xml.ModificacionDeATR.Cliente;

            var sb = new StringBuilder();
            sb.Append("REPLACE INTO cnmc_t_cliente ");
            sb.Append("(id, TipoIdentificador, Identificador, TipoPersona, NombreDePila, ");
            sb.Append("PrimerApellido, SegundoApellido, RazonSocial, created_by, created_date) VALUES (");

            sb.Append(id).Append(",");
            sb.Append(cli.IdCliente?.TipoIdentificador != null ? $"'{cli.IdCliente.TipoIdentificador}'," : "null,");
            sb.Append(cli.IdCliente?.Identificador != null ? $"'{cli.IdCliente.Identificador}'," : "null,");
            sb.Append(cli.IdCliente?.TipoPersona != null ? $"'{cli.IdCliente.TipoPersona}'," : "null,");
            sb.Append(cli.Nombre?.NombreDePila != null ? $"'{cli.Nombre.NombreDePila}'," : "null,");
            sb.Append(cli.Nombre?.PrimerApellido != null ? $"'{cli.Nombre.PrimerApellido}'," : "null,");
            sb.Append(cli.Nombre?.SegundoApellido != null ? $"'{cli.Nombre.SegundoApellido}'," : "null,");
            sb.Append(cli.Nombre?.RazonSocial != null ? $"'{cli.Nombre.RazonSocial}'," : "null,");
            sb.Append($"'{Environment.UserName.ToUpper()}',");
            sb.Append($"'{DateTime.Now:yyyy-MM-dd HH:mm:ss}')");

            var db = new MySQLDB(MySQLDB.Esquemas.CON);
            var command = new MySqlCommand(sb.ToString(), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }


        //
        private void Guarda_Contacto(int id, EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;

            if(xml.Alta.Contrato.Contacto.PersonaDeContacto != null)
            {
                sb.Append("REPLACE INTO cnmc_t_contacto");
                sb.Append(" (id, PersonaDeContacto, PrefijoPais, Numero, created_by, created_date) values ");
                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.Alta.Contrato.Contacto.PersonaDeContacto).Append("',");

                if (xml.Alta.Contrato.Contacto.Telefono.PrefijoPais != null)
                    sb.Append("'").Append(xml.Alta.Contrato.Contacto.Telefono.PrefijoPais).Append("',");
                else
                    sb.Append("null,");

                if (xml.Alta.Contrato.Contacto.Telefono.Numero != null)
                    sb.Append("'").Append(xml.Alta.Contrato.Contacto.Telefono.Numero).Append("',");
                else
                    sb.Append("null,");

                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }

            
        }
        private void Guarda_ContactoV30(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
           // bool firstOnly = true;

            if (xml.Alta.Contrato.Contacto.PersonaDeContacto != null)
            {
                sb.Append("REPLACE INTO cnmc_t_contacto");
                sb.Append(" (id, PersonaDeContacto, PrefijoPais, Numero, created_by, created_date) values ");
                sb.Append("(").Append(id).Append(",");
                sb.Append("'").Append(xml.Alta.Contrato.Contacto.PersonaDeContacto).Append("',");

                if (xml.Alta.Contrato.Contacto.Telefono.PrefijoPais != null)
                    sb.Append("'").Append(xml.Alta.Contrato.Contacto.Telefono.PrefijoPais).Append("',");
                else
                    sb.Append("null,");

                if (xml.Alta.Contrato.Contacto.Telefono.Numero != null)
                    sb.Append("'").Append(xml.Alta.Contrato.Contacto.Telefono.Numero).Append("',");
                else
                    sb.Append("null,");

                sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
                sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }


        }

        private void Guarda_Cliente(int id, EndesaEntity.cnmc.V21_2019_12_17.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();

            sb.Append("REPLACE INTO cnmc_t_cliente");
            sb.Append(" (id, TipoIdentificador, Identificador, TipoPersona, NombreDePila,");
            sb.Append(" PrimerApellido, SegundoApellido, RazonSocial,");
            sb.Append(" created_by, created_date) values ");

            sb.Append("(").Append(id).Append(",");
            
            if(xml.Alta.Cliente.IdCliente.TipoIdentificador != null)
                sb.Append("'").Append(xml.Alta.Cliente.IdCliente.TipoIdentificador).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.IdCliente.Identificador != null)
                sb.Append("'").Append(xml.Alta.Cliente.IdCliente.Identificador).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.IdCliente.TipoPersona != null)
                sb.Append("'").Append(xml.Alta.Cliente.IdCliente.TipoPersona).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.NombreDePila != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.NombreDePila).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.PrimerApellido != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.PrimerApellido).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.SegundoApellido != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.SegundoApellido).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.RazonSocial != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.RazonSocial).Append("',");
            else
                sb.Append("null,");

            sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
            sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        private void Guarda_ClienteV30(int id, EndesaEntity.cnmc.V30_2022_21_01.TipoMensajeA301 xml)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();

            sb.Append("REPLACE INTO cnmc_t_cliente");
            sb.Append(" (id, TipoIdentificador, Identificador, TipoPersona, NombreDePila,");
            sb.Append(" PrimerApellido, SegundoApellido, RazonSocial,");
            sb.Append(" created_by, created_date) values ");

            sb.Append("(").Append(id).Append(",");

            if (xml.Alta.Cliente.IdCliente.TipoIdentificador != null)
                sb.Append("'").Append(xml.Alta.Cliente.IdCliente.TipoIdentificador).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.IdCliente.Identificador != null)
                sb.Append("'").Append(xml.Alta.Cliente.IdCliente.Identificador).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.IdCliente.TipoPersona != null)
                sb.Append("'").Append(xml.Alta.Cliente.IdCliente.TipoPersona).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.NombreDePila != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.NombreDePila).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.PrimerApellido != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.PrimerApellido).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.SegundoApellido != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.SegundoApellido).Append("',");
            else
                sb.Append("null,");

            if (xml.Alta.Cliente.Nombre.RazonSocial != null)
                sb.Append("'").Append(xml.Alta.Cliente.Nombre.RazonSocial).Append("',");
            else
                sb.Append("null,");

            sb.Append("'").Append(System.Environment.UserName.ToUpper()).Append("',");
            sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
        private int GetID(string proceso, string paso, string codigo_solicitud)
        {
            int id = 0;
            List<EndesaEntity.xml.CNMC_t_cabecera> o;
            if (dic_cabeceras.TryGetValue(proceso + paso, out o))
            {
                EndesaEntity.xml.CNMC_t_cabecera c =
                    o.Find(z => z.CodigoDeSolicitud == codigo_solicitud);

                if(c != null)
                    return c.id;
                else
                    return GetLastID();
            }
            else
            {
                return GetLastID();
            }

            return id;
        }

        private int GetLastID()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            int id = 0;

            strSql = "select max(id) id from cnmc_t_cabecera";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                id = Convert.ToInt32(r["id"]);
            
            db.CloseConnection();

            return id + 1;

        }

        public void CreaMensajeA305(TipoMensajeA305 xml_a305)
        {
            throw new NotImplementedException();
        }
    }
}
