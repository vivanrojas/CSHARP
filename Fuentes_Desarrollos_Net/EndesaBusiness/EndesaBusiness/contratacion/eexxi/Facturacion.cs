using EndesaBusiness.servidores;
using EndesaEntity;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class Facturacion
    {
        public Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> dic_informe;
        EndesaBusiness.utilidades.Param p;
        public Facturacion()
        {
            dic_informe = Carga();
            p = new EndesaBusiness.utilidades.Param("eexxi_param", EndesaBusiness.servidores.MySQLDB.Esquemas.CON);
        }

        private Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> Carga()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;            

            Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> d =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion>();

            Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> dd =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion>();

            Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> ddd =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion>();

            Dictionary<string, string> dic_cups = new Dictionary<string, string>();

            EndesaBusiness.contratacion.PS_AT_HIST ps_at_hist_70;
            EndesaBusiness.contratacion.PS_AT_HIST ps_at_hist_20;

            DateTime primera_inserccion_PS_AT_HIST = new DateTime();

            try
            {
                // Averiguamos la primera anexion en PS_AT_HIST
                // para no buscar solicitudes anteriores a esa fecha

                strSql = "SELECT min(h.Fecha_Anexion) AS min_fecha FROM PS_AT_HIST h";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["min_fecha"] != System.DBNull.Value)
                        primera_inserccion_PS_AT_HIST = Convert.ToDateTime(r["min_fecha"]);
                }
                db.CloseConnection();

                #region Consulta ALTAS Y BAJAS eexxi_solicitudes_tmp
                strSql = "SELECT sol.CodigoDeSolicitud,"
                    + " sol.CUPS, sol.Identificador AS NIF,"
                    + " sol.RazonSocial AS CLIENTE,"
                    + " sol.FechaActivacion AS FECHA_ALTA,"                    
                    + " bajas.FechaActivacion as FECHA_BAJA,"
                    + " DATE_FORMAT(CONCAT(substr(ps.fAltaCont,1,4),substr(ps.fAltaCont,5,2),substr(ps.fAltaCont,7,2)),'%Y-%m-%d') AS ALTA_EE,"
                    + " DATEDIFF(bajas.FechaActivacion, sol.FechaActivacion) as DIF,"
                    + " (sol.PotenciaPeriodo1 / 1000) AS P1,"
                    + " (sol.PotenciaPeriodo2 / 1000) AS P2,"
                    + " (sol.PotenciaPeriodo3 / 1000) AS P3,"
                    + " (sol.PotenciaPeriodo4 / 1000) AS P4,"
                    + " (sol.PotenciaPeriodo5 / 1000) AS P5,"
                    + " (sol.PotenciaPeriodo6 / 1000) AS P6,"
                    + " atr.descripcion AS TARIFA,"
                    + " ten.descripcion AS TENSION,"
                    + " NULL, NULL, NOW()"
                    + " FROM cont.eexxi_solicitudes_tmp sol"
                    + " INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + " c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + " c.CodigoDePaso = sol.CodigoDePaso"
                    + " LEFT OUTER JOIN eexxi_para_facturar f ON"
                    + " f.cups = sol.CUPS"
                    + " AND f.fecha_alta = sol.FechaActivacion"
                    + " INNER JOIN(SELECT CUPS, MAX(FechaActivacion) FechaActivacion FROM"
                    + "     cont.eexxi_solicitudes_tmp sol"
                    + "     INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + "     c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + "     c.CodigoDePaso = sol.CodigoDePaso"
                    + "     WHERE"
                    + "     c.Descripcion = 'ALTA'"
                    + "     GROUP BY CUPS) AS altas on"
                    + " altas.CUPS = sol.CUPS"
                    + " AND altas.FechaActivacion = sol.FechaActivacion"
                    + " LEFT OUTER JOIN(SELECT CUPS, MAX(FechaActivacion) FechaActivacion FROM"
                    + "     cont.eexxi_solicitudes_tmp sol"
                    + "     INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + "     c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + "     c.CodigoDePaso = sol.CodigoDePaso"
                    + "     WHERE"
                    + "     c.Descripcion = 'BAJA'"
                    + "     GROUP BY CUPS) AS bajas on"
                    + " bajas.CUPS = sol.CUPS"                   
                    + " INNER JOIN cont.eexxi_param_codigos_tarifas_atr atr ON"
                    + " atr.tarifa_atr = sol.TarifaATR"
                    + " INNER JOIN cont.eexxi_param_codigos_tensiones ten ON"
                    + " ten.codigo = sol.TensionDelSuministro"
                    + " INNER JOIN cont.PS_AT ps ON"
                    + " ps.CUPS22 = sol.CUPS AND"
                    + " ps.EMPRESA = 'EE'"
                    + " WHERE c.Descripcion = 'ALTA'"                          
                    + " AND (sol.FechaActivacion <= bajas.FechaActivacion)"
                    + " AND f.codigodesolicitud IS NULL"
                    + " ORDER BY sol.FechaActivacion DESC";
                #endregion
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Informe_Facturacion c = new EndesaEntity.contratacion.xxi.Informe_Facturacion();

                    if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                        c.codigo_solicitud = r["CodigoDeSolicitud"].ToString();

                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups = r["CUPS"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["FECHA_ALTA"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["FECHA_ALTA"]);

                    if (r["FECHA_BAJA"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["FECHA_BAJA"]);

                    if (r["ALTA_EE"] != System.DBNull.Value)
                        c.alta_ee = Convert.ToDateTime(r["ALTA_EE"]);

                    if (r["P1"] != System.DBNull.Value)
                        c.p1 = Convert.ToDouble(r["P1"]);

                    if (r["P2"] != System.DBNull.Value)
                        c.p2 = Convert.ToDouble(r["P2"]);

                    if (r["P3"] != System.DBNull.Value)
                        c.p3 = Convert.ToDouble(r["P3"]);

                    if (r["P4"] != System.DBNull.Value)
                        c.p4 = Convert.ToDouble(r["P4"]);

                    if (r["P5"] != System.DBNull.Value)
                        c.p5 = Convert.ToDouble(r["P5"]);

                    if (r["P6"] != System.DBNull.Value)
                        c.p6 = Convert.ToDouble(r["P6"]);

                    if (r["TARIFA"] != System.DBNull.Value)
                        c.tarifa = r["TARIFA"].ToString();

                    if (r["TENSION"] != System.DBNull.Value)
                        c.tension = Convert.ToDouble(r["TENSION"]);

                    if(c.fecha_baja > primera_inserccion_PS_AT_HIST)
                    {
                        EndesaEntity.contratacion.xxi.Informe_Facturacion o;
                        if (!d.TryGetValue(c.cups, out o))
                            d.Add(c.cups, c);

                        string oo;
                        if (!dic_cups.TryGetValue(c.cups, out oo))
                            dic_cups.Add(c.cups, c.cups);
                    }

                    

                }
                db.CloseConnection();

                #region Consulta ALTAS eexxi_solicitudes Y BAJAS eexxi_solicitudes_tmp
                strSql = "SELECT sol.CodigoDeSolicitud,"
                    + " sol.CUPS, sol.Identificador AS NIF,"
                    + " sol.RazonSocial AS CLIENTE,"
                    + " sol.FechaActivacion AS FECHA_ALTA,"                    
                    + " bajas.FechaActivacion as FECHA_BAJA,"
                    + " DATE_FORMAT(CONCAT(substr(ps.fAltaCont,1,4),substr(ps.fAltaCont,5,2),substr(ps.fAltaCont,7,2)),'%Y-%m-%d') AS ALTA_EE,"
                    + " DATEDIFF(bajas.FechaActivacion, sol.FechaActivacion) as DIF,"
                    + " (sol.PotenciaPeriodo1 / 1000) AS P1,"
                    + " (sol.PotenciaPeriodo2 / 1000) AS P2,"
                    + " (sol.PotenciaPeriodo3 / 1000) AS P3,"
                    + " (sol.PotenciaPeriodo4 / 1000) AS P4,"
                    + " (sol.PotenciaPeriodo5 / 1000) AS P5,"
                    + " (sol.PotenciaPeriodo6 / 1000) AS P6,"
                    + " atr.descripcion AS TARIFA,"
                    + " ten.descripcion AS TENSION,"
                    + " NULL, NULL, NOW(), f.f_envio_informe"
                    + " FROM cont.eexxi_solicitudes sol"
                    + " INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + " c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + " c.CodigoDePaso = sol.CodigoDePaso"
                    + " LEFT OUTER JOIN eexxi_para_facturar f ON"
                    + " f.cups = sol.CUPS"
                    + " AND f.fecha_alta = sol.FechaActivacion"
                    + " INNER JOIN(SELECT CUPS, MAX(FechaActivacion) FechaActivacion FROM"
                    + "     cont.eexxi_solicitudes sol"
                    + "     INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + "     c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + "     c.CodigoDePaso = sol.CodigoDePaso"
                    + "     WHERE"
                    + "     c.Descripcion = 'ALTA'"
                    + "     GROUP BY CUPS) AS altas on"
                    + " altas.CUPS = sol.CUPS"
                    + " AND altas.FechaActivacion = sol.FechaActivacion"
                    + " LEFT OUTER JOIN (SELECT CUPS, MAX(FechaActivacion) FechaActivacion FROM"
                    + "     cont.eexxi_solicitudes_tmp sol" 
                    + "     INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + "     c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + "     c.CodigoDePaso = sol.CodigoDePaso"
                    + "     WHERE"
                    + "     c.Descripcion = 'BAJA'"
                    + "     GROUP BY CUPS) AS bajas on"
                    + " bajas.CUPS = sol.CUPS"                    
                    + " INNER JOIN cont.eexxi_param_codigos_tarifas_atr atr ON"
                    + " atr.tarifa_atr = sol.TarifaATR"
                    + " INNER JOIN cont.eexxi_param_codigos_tensiones ten ON"
                    + " ten.codigo = sol.TensionDelSuministro"
                    + " INNER JOIN cont.PS_AT ps ON"
                    + " ps.CUPS22 = sol.CUPS AND"
                    + " ps.EMPRESA = 'EE'"
                    + " WHERE c.Descripcion = 'ALTA'"                                    
                    + " AND (sol.FechaActivacion <= bajas.FechaActivacion)"
                    + " AND f.codigodesolicitud IS NULL"
                    + " ORDER BY sol.FechaActivacion DESC";
                #endregion
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Informe_Facturacion c = new EndesaEntity.contratacion.xxi.Informe_Facturacion();

                    if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                        c.codigo_solicitud = r["CodigoDeSolicitud"].ToString();

                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups = r["CUPS"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["FECHA_ALTA"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["FECHA_ALTA"]);

                    if (r["FECHA_BAJA"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["FECHA_BAJA"]);

                    if (r["ALTA_EE"] != System.DBNull.Value)
                        c.alta_ee = Convert.ToDateTime(r["ALTA_EE"]);

                    if (r["P1"] != System.DBNull.Value)
                        c.p1 = Convert.ToDouble(r["P1"]);

                    if (r["P2"] != System.DBNull.Value)
                        c.p2 = Convert.ToDouble(r["P2"]);

                    if (r["P3"] != System.DBNull.Value)
                        c.p3 = Convert.ToDouble(r["P3"]);

                    if (r["P4"] != System.DBNull.Value)
                        c.p4 = Convert.ToDouble(r["P4"]);

                    if (r["P5"] != System.DBNull.Value)
                        c.p5 = Convert.ToDouble(r["P5"]);

                    if (r["P6"] != System.DBNull.Value)
                        c.p6 = Convert.ToDouble(r["P6"]);

                    if (r["TARIFA"] != System.DBNull.Value)
                        c.tarifa = r["TARIFA"].ToString();

                    if (r["TENSION"] != System.DBNull.Value)
                        c.tension = Convert.ToDouble(r["TENSION"]);

                    if (r["f_envio_informe"] != System.DBNull.Value)
                        c.fecha_envio_mail = Convert.ToDateTime(r["f_envio_informe"]);
                    //else
                    //{
                    //    if (Convert.ToDateTime(r["last_update_date"]) <
                    //        new DateTime(2022, 08, 01))
                    //        c.fecha_envio_mail = new DateTime(1901, 01, 01);
                    //}

                    if (c.fecha_baja > primera_inserccion_PS_AT_HIST)
                    {
                        EndesaEntity.contratacion.xxi.Informe_Facturacion o;
                        if (!d.TryGetValue(c.cups, out o))
                            d.Add(c.cups, c);

                        string oo;
                        if (!dic_cups.TryGetValue(c.cups, out oo))
                            dic_cups.Add(c.cups, c.cups);
                    }

                       

                }
                db.CloseConnection();

                #region Consulta ALTAS Y BAJAS eexxi_solicitudes
                strSql = "SELECT sol.CodigoDeSolicitud,"
                    + " sol.CUPS, sol.Identificador AS NIF,"
                    + " sol.RazonSocial AS CLIENTE,"
                    + " sol.FechaActivacion AS FECHA_ALTA,"                    
                    + " bajas.FechaActivacion as FECHA_BAJA,"
                    + " DATE_FORMAT(CONCAT(substr(ps.fAltaCont,1,4),substr(ps.fAltaCont,5,2),substr(ps.fAltaCont,7,2)),'%Y-%m-%d') AS ALTA_EE,"
                    + " DATEDIFF(bajas.FechaActivacion, sol.FechaActivacion) as DIF,"
                    + " (sol.PotenciaPeriodo1 / 1000) AS P1,"
                    + " (sol.PotenciaPeriodo2 / 1000) AS P2,"
                    + " (sol.PotenciaPeriodo3 / 1000) AS P3,"
                    + " (sol.PotenciaPeriodo4 / 1000) AS P4,"
                    + " (sol.PotenciaPeriodo5 / 1000) AS P5,"
                    + " (sol.PotenciaPeriodo6 / 1000) AS P6,"
                    + " atr.descripcion AS TARIFA,"
                    + " ten.descripcion AS TENSION,"
                    + " NULL, NULL, NOW(), f.f_envio_informe"
                    + " FROM cont.eexxi_solicitudes sol"
                    + " INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + " c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + " c.CodigoDePaso = sol.CodigoDePaso"
                    + " LEFT OUTER JOIN eexxi_para_facturar f ON"
                    + " f.cups = sol.CUPS"
                    + " AND f.fecha_alta = sol.FechaActivacion"
                    + " INNER JOIN (SELECT CUPS, MAX(FechaActivacion) FechaActivacion FROM"
                    + "     cont.eexxi_solicitudes sol"
                    + "     INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + "     c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + "     c.CodigoDePaso = sol.CodigoDePaso"
                    + "     WHERE"
                    + "     c.Descripcion = 'ALTA'"
                    + "     GROUP BY CUPS) AS altas on"
                    + " altas.CUPS = sol.CUPS"
                    + " AND altas.FechaActivacion = sol.FechaActivacion"
                    + " LEFT OUTER JOIN(SELECT CUPS, MAX(FechaActivacion) FechaActivacion FROM"
                    + "     cont.eexxi_solicitudes sol"
                    + "     INNER JOIN eexxi_param_solicitudes_codigos c ON"
                    + "     c.CodigoDelProceso = sol.CodigoDelProceso AND"
                    + "     c.CodigoDePaso = sol.CodigoDePaso"
                    + "     WHERE"
                    + "     c.Descripcion = 'BAJA'"
                    + "     GROUP BY CUPS) AS bajas on"
                    + " bajas.CUPS = sol.CUPS"                    
                    + " INNER JOIN cont.eexxi_param_codigos_tarifas_atr atr ON"
                    + " atr.tarifa_atr = sol.TarifaATR"
                    + " INNER JOIN cont.eexxi_param_codigos_tensiones ten ON"
                    + " ten.codigo = sol.TensionDelSuministro"
                    + " INNER JOIN cont.PS_AT ps ON"
                    + " ps.CUPS22 = sol.CUPS AND"
                    + " ps.EMPRESA = 'EE'"
                    + " WHERE c.Descripcion = 'ALTA'"                             
                    + " AND (sol.FechaActivacion <= bajas.FechaActivacion)"
                    + " AND f.codigodesolicitud IS NULL"
                    + " ORDER BY sol.FechaActivacion DESC";
                #endregion
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Informe_Facturacion c = new EndesaEntity.contratacion.xxi.Informe_Facturacion();

                    if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                        c.codigo_solicitud = r["CodigoDeSolicitud"].ToString();

                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups = r["CUPS"].ToString();

                    if (r["NIF"] != System.DBNull.Value)
                        c.nif = r["NIF"].ToString();

                    if (r["CLIENTE"] != System.DBNull.Value)
                        c.cliente = r["CLIENTE"].ToString();

                    if (r["FECHA_ALTA"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["FECHA_ALTA"]);

                    if (r["FECHA_BAJA"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["FECHA_BAJA"]);

                    if (r["ALTA_EE"] != System.DBNull.Value)
                        c.alta_ee = Convert.ToDateTime(r["ALTA_EE"]);

                    if (r["DIF"] != System.DBNull.Value)
                        c.dif = Convert.ToInt32(r["DIF"]);

                    if (r["P1"] != System.DBNull.Value)
                        c.p1 = Convert.ToDouble(r["P1"]);

                    if (r["P2"] != System.DBNull.Value)
                        c.p2 = Convert.ToDouble(r["P2"]);

                    if (r["P3"] != System.DBNull.Value)
                        c.p3 = Convert.ToDouble(r["P3"]);

                    if (r["P4"] != System.DBNull.Value)
                        c.p4 = Convert.ToDouble(r["P4"]);

                    if (r["P5"] != System.DBNull.Value)
                        c.p5 = Convert.ToDouble(r["P5"]);

                    if (r["P6"] != System.DBNull.Value)
                        c.p6 = Convert.ToDouble(r["P6"]);

                    if (r["TARIFA"] != System.DBNull.Value)
                        c.tarifa = r["TARIFA"].ToString();

                    if (r["TENSION"] != System.DBNull.Value)
                        c.tension = Convert.ToDouble(r["TENSION"]);

                    if (r["f_envio_informe"] != System.DBNull.Value)
                        c.fecha_envio_mail = Convert.ToDateTime(r["f_envio_informe"]);
                    //else
                    //{
                    //    if (Convert.ToDateTime(r["last_update_date"]) <
                    //        new DateTime(2022, 08, 01))
                    //        c.fecha_envio_mail = new DateTime(1901, 01, 01);
                    //}

                    if (c.fecha_baja > primera_inserccion_PS_AT_HIST)
                    {
                        EndesaEntity.contratacion.xxi.Informe_Facturacion o;
                        if (!d.TryGetValue(c.cups, out o))
                            d.Add(c.cups, c);

                        string oo;
                        if (!dic_cups.TryGetValue(c.cups, out oo))
                            dic_cups.Add(c.cups, c.cups);
                    }

                       

                }
                db.CloseConnection();

                // Añadimos el resto del histórico

                strSql = "select codigodesolicitud, cups, nif, cliente,"
                    + " fecha_alta, fecha_baja, alta_ee, dif,"
                    + " p1, p2, p3, p4, p5, p6, tarifa, tension, usuario, f_envio_informe"
                    + " from eexxi_para_facturar";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Informe_Facturacion c = 
                        new EndesaEntity.contratacion.xxi.Informe_Facturacion();
                    if (r["codigodesolicitud"] != System.DBNull.Value)
                        c.codigo_solicitud = r["codigodesolicitud"].ToString();

                    if (r["cups"] != System.DBNull.Value)
                        c.cups = r["cups"].ToString();

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["cliente"] != System.DBNull.Value)
                        c.cliente = r["cliente"].ToString();

                    if (r["fecha_alta"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["fecha_alta"]);

                    if (r["fecha_baja"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["fecha_baja"]);

                    if (r["alta_ee"] != System.DBNull.Value)
                        c.alta_ee = Convert.ToDateTime(r["alta_ee"]);

                    if (r["dif"] != System.DBNull.Value)
                        c.dif = Convert.ToInt32(r["dif"]);

                    if (r["p1"] != System.DBNull.Value)
                        c.p1 = Convert.ToDouble(r["p1"]);

                    if (r["p2"] != System.DBNull.Value)
                        c.p2 = Convert.ToDouble(r["p2"]);

                    if (r["p3"] != System.DBNull.Value)
                        c.p3 = Convert.ToDouble(r["p3"]);

                    if (r["p4"] != System.DBNull.Value)
                        c.p4 = Convert.ToDouble(r["p4"]);

                    if (r["p5"] != System.DBNull.Value)
                        c.p5 = Convert.ToDouble(r["p5"]);

                    if (r["p6"] != System.DBNull.Value)
                        c.p6 = Convert.ToDouble(r["p6"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();

                    if (r["tension"] != System.DBNull.Value)
                        c.tension = Convert.ToDouble(r["tension"]);

                    if (r["f_envio_informe"] != System.DBNull.Value)
                        c.fecha_envio_mail = Convert.ToDateTime(r["f_envio_informe"]);
                    else
                    {
                        if (Convert.ToDateTime(r["last_update_date"]) < 
                            new DateTime(2022,08, 01))
                        c.fecha_envio_mail = new DateTime(1901, 01, 01);
                    }

                    if (c.fecha_baja > primera_inserccion_PS_AT_HIST)
                    {
                        EndesaEntity.contratacion.xxi.Informe_Facturacion o;
                        if (!d.TryGetValue(c.cups, out o))
                            d.Add(c.cups, c);

                        string oo;
                        if (!dic_cups.TryGetValue(c.cups, out oo))
                            dic_cups.Add(c.cups, c.cups);
                    }

                       

                }
                db.CloseConnection();


                // 1.- Que tenga solicitud de alta para el Empresa 70
                // 2.- Que tenga solicitud de baja para la empresa 70
                // 3.- Que no tenga contrato activo (PS_AT_HIST) para la empresa 70
                // 4.- Que tenga contrato activo (PS_AT) para le empresa 20
                // 5.- La fecha de inicio del contrato sea 1 día posterior de la fecha de baja
                // de la empresa 70.

                EndesaBusiness.contratacion.SolicitudesATR sol
                     = new SolicitudesATR(dic_cups.Values.ToList(), 70, "ACTIVADA", "BAJA");
                DateTime f = new DateTime();

                // Buscamos en PS_AT_HIST que no esté dado de alta en XXI


                ps_at_hist_70 = new PS_AT_HIST(dic_cups.Values.ToList(), "EEXXI");
                ps_at_hist_20 = new PS_AT_HIST(dic_cups.Values.ToList(), "EE");



                //foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> p in d)
                //{
                //    if (!sol.ExisteCUPS(p.Key))
                //        dd.Add(p.Key, p.Value);
                //    else
                //    {
                //        string dia = Convert.ToString(sol.fAcepRech);
                //        f = new DateTime(Convert.ToInt32(dia.Substring(0, 4)),
                //            Convert.ToInt32(dia.Substring(4, 2)), Convert.ToInt32(dia.Substring(6, 2)));
                //        if (f < p.Value.fecha_alta)
                //            dd.Add(p.Key, p.Value);
                //    }
                //}

                foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> p in d)
                {
                    if (AddCUPS_EE(ps_at_hist_70, ps_at_hist_20,
                        //p.Value.cups.Substring(0, 20), p.Value.fecha_alta, p.Value.fecha_baja))
                        p.Value.cups, p.Value.fecha_alta, p.Value.fecha_baja))
                    {
                        // Añadimos verificación de si CUPS|FECHA_ALTA se encuentra en tabla eexxi_no_facturar
                        strSql = "SELECT COUNT(*) AS contador FROM eexxi_no_facturar f WHERE f.cups = '" +
                                p.Value.cups + "' AND f.fecha_alta ='" + p.Value.fecha_alta.ToString("yyyy-MM-dd") + "';";
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(strSql, db.con);
                        r = command.ExecuteReader();
                        while (r.Read())
                        {
                            if (Convert.ToInt32(r["contador"]) == 0)
                                ddd.Add(p.Key, p.Value);
                        }
                        db.CloseConnection();
                        
                    }


                }
                               

                return ddd;

            }catch(Exception ex)
            {

                MessageBox.Show(ex.Message,
                  "eexxi.Facturacion.Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);

                return null;
            }
        }


        private bool AddCUPS_EE(EndesaBusiness.contratacion.PS_AT_HIST ps_at_hist_70,
            EndesaBusiness.contratacion.PS_AT_HIST ps_at_hist_20,
            string cups20, DateTime fecha_alta, DateTime fecha_baja)
        {
            bool existe_70 = false;
            bool existe_20 = false;
            

            // Que no tenga contrato activo en EEXXI 

            List<EndesaEntity.contratacion.PS_AT_Tabla> o;
            if (ps_at_hist_70.dic.TryGetValue(cups20, out o))
            {
                foreach(EndesaEntity.contratacion.PS_AT_Tabla p in o)
                {
                    existe_70 = p.fecha_alta_contrato == fecha_alta;
                    if (existe_70)
                        break;
                }
            }
                
               

            if (!existe_70)
            {
                List<EndesaEntity.contratacion.PS_AT_Tabla> oo;
                if (ps_at_hist_20.dic.TryGetValue(cups20, out oo))
                {
                    // Quitamos la condición que:
                    // 5.- La fecha de inicio del contrato sea 1 día posterior de la fecha de baja
                    // de la empresa 70.
                    // Si existe el contrato en vigor en la empresa
                    //existe_20 = fecha_baja.AddDays(1) == oo.fecha_alta_contrato;

                    foreach (EndesaEntity.contratacion.PS_AT_Tabla p in oo)
                    {
                        existe_20 = fecha_baja < p.fecha_alta_contrato;
                        if (existe_20)
                            break;
                    }
                    
                }
            }
                        return !existe_70 && existe_20;

        }



        public void InformeViaMail(string rutaSalida, bool enviarMail, Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> dic)
        {
            
            int f = 0; // fila
            int c = 0; // columna
            StringBuilder cuerpo = new StringBuilder();
            EndesaBusiness.office.MailCompose mail = new EndesaBusiness.office.MailCompose();

            try
            {
                FileInfo file = new FileInfo(rutaSalida);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(file);
                var workSheet = excelPackage.Workbook.Worksheets.Add("ALTAS EEXXI");
                var headerCells = workSheet.Cells[1, 1, 1, 13];
                var headerFont = headerCells.Style.Font;

                var allCells = workSheet.Cells[1, 1, 1, 13];
                var cellFont = allCells.Style.Font;
                cellFont.Bold = true;

                f = 1;
                c = 1;

                workSheet.Cells[f, c].Value = "CUPS";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "NIF";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "CLIENTE";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "FECHA DE ALTA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "FECHA DE BAJA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                for (int i = 1; i <= 6; i++)
                {
                    workSheet.Cells[f, c].Value = "P" + i;
                    workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;
                }

                workSheet.Cells[f, c].Value = "TARIFA";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                workSheet.Cells[f, c].Value = "TENSIÓN";
                workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                foreach(KeyValuePair<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> p in dic)
                {
                    f++;
                    c = 1;
                    workSheet.Cells[f, c].Value = p.Value.cups; c++;
                    workSheet.Cells[f, c].Value = p.Value.nif; c++;
                    workSheet.Cells[f, c].Value = p.Value.cliente; c++;
                    workSheet.Cells[f, c].Value = p.Value.fecha_alta;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                    workSheet.Cells[f, c].Value = p.Value.fecha_baja;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;

                    if(p.Value.p1 > 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.p1;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (p.Value.p2 > 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.p2;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (p.Value.p3 > 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.p3;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (p.Value.p4 > 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.p4;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (p.Value.p5 > 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.p5;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    if (p.Value.p6 > 0)
                    {
                        workSheet.Cells[f, c].Value = p.Value.p6;
                        workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                    }
                    c++;

                    workSheet.Cells[f, c].Value = p.Value.tarifa; c++;
                    workSheet.Cells[f, c].Value = p.Value.tension;
                    workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0"; c++;
                }

                allCells = workSheet.Cells[1, 1, f, c];

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:M1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();
                excelPackage = null;

                if (enviarMail)
                {
                    cuerpo.Append(System.Environment.NewLine);
                    cuerpo.Append(DateTime.Now.Hour > 14 ? "Buenas tardes:" : "Buenos días:");
                    cuerpo.Append(System.Environment.NewLine);
                    cuerpo.Append("Se adjunta el Excel de los puntos de EEXXI para facturar por varios. ");
                    cuerpo.Append(System.Environment.NewLine);
                    cuerpo.Append("Un saludo.");

                    mail.para.Add(p.GetValue("InformeFacturasPara"));
                    //mail.cc.Add(p.GetValue("MailCurvasCC"));
                    mail.asunto = "";
                    mail.htmlCuerpo = cuerpo.ToString();
                    mail.adjuntos.Add(file.FullName);
                    mail.Show();
                    file.Delete();

                    GuardarDatos(dic);
                }


                

            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                  "InformeViaMail",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }

        }

        private void GuardarDatos(Dictionary<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> dic)
        {
            
            MySQLDB db;
            MySqlCommand command;
            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            int x = 0;
            int i = 0;

            try
            {
                foreach(KeyValuePair<string, EndesaEntity.contratacion.xxi.Informe_Facturacion> p in dic)
                {
                    x++;
                    i++;

                    if (firstOnly)
                    {
                        sb.Append("replace INTO eexxi_para_facturar (codigodesolicitud, cups, nif, cliente, fecha_alta,");
                        sb.Append("fecha_baja, alta_ee, dif, p1, p2, p3, p4, p5, p6, tarifa, tension,");
                        sb.Append("usuario, f_envio_informe, f_ult_mod) values ");

                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.Value.codigo_solicitud).Append("',");
                    sb.Append("'").Append(p.Value.cups).Append("',");
                    sb.Append("'").Append(p.Value.nif).Append("',");
                    sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(p.Value.cliente)).Append("',");

                    if(p.Value.fecha_alta > DateTime.MinValue)
                        sb.Append("'").Append(p.Value.fecha_alta.ToString("yyyy-MM-dd")).Append("',");

                    if (p.Value.fecha_baja > DateTime.MinValue)
                        sb.Append("'").Append(p.Value.fecha_baja.ToString("yyyy-MM-dd")).Append("',");

                    if (p.Value.alta_ee > DateTime.MinValue)
                        sb.Append("'").Append(p.Value.alta_ee.ToString("yyyy-MM-dd")).Append("',");

                    sb.Append(p.Value.dif).Append(",");

                    if (p.Value.p1 > 0)
                        sb.Append(p.Value.p1.ToString().Replace(",",".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.Value.p2 > 0)
                        sb.Append(p.Value.p2.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.Value.p3 > 0)
                        sb.Append(p.Value.p3.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.Value.p4 > 0)
                        sb.Append(p.Value.p4.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.Value.p5 > 0)
                        sb.Append(p.Value.p5.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    if (p.Value.p6 > 0)
                        sb.Append(p.Value.p6.ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");

                    sb.Append("'").Append(p.Value.tarifa).Append("',");
                    sb.Append(p.Value.tension).Append(",");
                    sb.Append("'").Append(System.Environment.UserName).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                    if (x == 250)
                    {

                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        x = 0;
                    }
                }

                if (x > 0)
                {

                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    x = 0;
                }


            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                   "GuardarDatos",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }

            
        }
        public void AddNoFacturar(EndesaEntity.contratacion.xxi.Informe_Facturacion o)
        {

            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            
            try
            {
                sb.Append("replace INTO eexxi_no_facturar (cups, nif, cliente, fecha_alta,");
                sb.Append("fecha_baja, created_by) values (");
                sb.Append("'").Append(o.cups).Append("',");
                sb.Append("'").Append(o.nif).Append("',");
                sb.Append("'").Append(utilidades.FuncionesTexto.ArreglaAcentos(o.cliente)).Append("',");
                if (o.fecha_alta > DateTime.MinValue)
                    sb.Append("'").Append(o.fecha_alta.ToString("yyyy-MM-dd")).Append("',");

                if (o.fecha_baja > DateTime.MinValue)
                    sb.Append("'").Append(o.fecha_baja.ToString("yyyy-MM-dd")).Append("',");

                sb.Append("'").Append(System.Environment.UserName).Append("');");

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "AddNoFacturar",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }


        }

    }
}

