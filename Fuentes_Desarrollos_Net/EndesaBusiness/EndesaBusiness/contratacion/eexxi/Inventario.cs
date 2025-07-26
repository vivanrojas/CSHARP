using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
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
    public class Inventario : EndesaEntity.contratacion.Inventario_Tabla
    {
        public Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla> dic { get; set; }
        public Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla> dic_tmp_altas { get; set; }
        public Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla> dic_tmp_bajas { get; set; }

        public Inventario()
        {


            dic = new Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla>();
            dic_tmp_altas = new Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla>();
            dic_tmp_bajas = new Dictionary<string, EndesaEntity.contratacion.Inventario_Tabla>();
            Carga_Solicitudes();
            CargaTemporal();
                                    

        }


        public void AnalizaInventario(string cups22, string estado)
        {
            EndesaEntity.contratacion.Inventario_Tabla o;
            if (dic.TryGetValue(cups22, out o))
            {
                //this.id_inventario = o.id_inventario;
                this.cups22 = o.cups22;

            }
            else if (dic_tmp_altas.TryGetValue(cups22, out o))
            {
                // this.id_inventario = o.id_inventario;
                this.cups22 = o.cups22;

            }
            else
            {
                EndesaEntity.contratacion.Inventario_Tabla c = new EndesaEntity.contratacion.Inventario_Tabla();
                // c.id_inventario = dic.Count + dic_tmp_altas.Count + 1;
                c.cups22 = cups22;
                c.estado = estado;
                // this.id_inventario = c.id_inventario;
                this.estado = c.estado;
                dic_tmp_altas.Add(c.cups22, c);
            }
        }


        public void Carga_Solicitudes()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {


                strSql = "select t101.Identificador as nif, t101.RazonSocial as nombre_cliente," 
                    + " s.CUPS, s.FechaActivacion, s.RazonSocial, s.Identificador, s.informado_alta,"
                    + " if (p.DesProvincia IS null , if (pp.DesProvincia IS null, ppp.DesProvincia, pp.DesProvincia), p.DesProvincia) as Provincia," 
                    + " s.DescripcionPoblacion, ec.es_incidencia";
                for (int i = 1; i < 7; i++)
                    strSql += " ,s.PotenciaPeriodo" + i;

                strSql += " from cont.eexxi_solicitudes s inner join"
                    + " cont.eexxi_param_solicitudes_codigos sc on"
                    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                    + " sc.CodigoDePaso = s.CodigoDePaso"
                    + " LEFT OUTER JOIN cont.eexxi_solicitudes_t101 t101 ON"
                    + " t101.CodigoDeSolicitud = s.CodigoDeSolicitud AND"
                    + " t101.CUPS = s.CUPS"
                    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                    //+ " h.cups = substr(s.CUPS, 1, 20)"                    
                    + " h.cups = s.CUPS"
                    // 20221129 ADD
                    + " and h.codigo_solicitud = s.CodigoDeSolicitud"
                    // 20221129 END ADD
                    + " LEFT OUTER JOIN eexxi_param_estados_casos ec ON"
                    + " ec.estado_id = h.estado_id"                    
                    + " left outer join eexxi_param_provincias p ON"
                    + " substr(s.CodPostalCliente, 1, 2) = p.CodigoPostal"
                    + " left outer join eexxi_param_provincias pp ON"
                    + " substr(s.CodPostal, 1, 2) = pp.CodigoPostal"
                    + " left outer join eexxi_param_provincias ppp ON"
                    + " substr(t101.CodPostalCliente, 1, 2) = ppp.CodigoPostal"
                    + " where sc.Descripcion = 'ALTA' order by s.FechaActivacion DESC";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Inventario_Tabla c = new EndesaEntity.contratacion.Inventario_Tabla();

                    c.cups22 = r["CUPS"].ToString();

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["nombre_cliente"] != System.DBNull.Value)
                        c.razon_social = r["nombre_cliente"].ToString();

                    if (r["es_incidencia"] != System.DBNull.Value)
                        c.tiene_incidencia = r["es_incidencia"].ToString() == "S";

                    if (r["FechaActivacion"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["FechaActivacion"]);

                    if (r["informado_alta"] != System.DBNull.Value)
                        c.informado_alta = Convert.ToDateTime(r["informado_alta"]);

                    if (r["RazonSocial"] != System.DBNull.Value)
                        if(c.razon_social == null)
                            c.razon_social = r["RazonSocial"].ToString();

                    if (r["Identificador"] != System.DBNull.Value)
                        if (c.nif == null)
                            c.nif = r["Identificador"].ToString();

                    if (r["Provincia"] != System.DBNull.Value)
                        c.cod_provincia = r["Provincia"].ToString();

                    if (r["DescripcionPoblacion"] != System.DBNull.Value)
                        c.descripcion_poblacion = r["DescripcionPoblacion"].ToString();

                    for (int i = 1; i < 7; i++)
                        if (r["PotenciaPeriodo" + i] != System.DBNull.Value)
                            c.potencias[i - 1] = Convert.ToDouble(r["PotenciaPeriodo" + i]);

                    c.vigente = true;

                    EndesaEntity.contratacion.Inventario_Tabla o;
                    //if (!dic.TryGetValue(c.cups22.Substring(0, 20), out o))
                    if (!dic.TryGetValue(c.cups22, out o))
                        //dic.Add(c.cups22.Substring(0, 20), c);
                        dic.Add(c.cups22, c);
                }
                db.CloseConnection();

                strSql = "select s.CUPS, s.FechaActivacion, s.informado_alta"
                    + " from cont.eexxi_solicitudes s inner join"
                    + " cont.eexxi_param_solicitudes_codigos sc on"
                    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                    + " sc.CodigoDePaso = s.CodigoDePaso"
                    + " where sc.Descripcion = 'BAJA'"
                    + " order by s.FechaActivacion desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Inventario_Tabla c = new EndesaEntity.contratacion.Inventario_Tabla();

                    c.cups22 = r["CUPS"].ToString();

                    if (r["FechaActivacion"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["FechaActivacion"]);

                    if (r["informado_alta"] != System.DBNull.Value)
                        c.informado_alta = Convert.ToDateTime(r["informado_alta"]);

                        EndesaEntity.contratacion.Inventario_Tabla o;
                    //if (dic.TryGetValue(c.cups22.Substring(0, 20), out o))
                    if (dic.TryGetValue(c.cups22, out o))
                    {
                        //if (o.fecha_alta <= c.fecha_baja)
                        if (o.fecha_alta <= c.fecha_baja.AddDays(1))
                        {
                            o.vigente = false;
                            o.fecha_baja = c.fecha_baja;
                            o.informado_alta = c.informado_alta;
                        }
                    }

                }
                db.CloseConnection();


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "EEXXI - CargaInventario",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public void CargaTemporal()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                #region ALTAS_TEMP
                strSql = "select t101.Identificador as nif, t101.RazonSocial as nombre_cliente,"
                    + " s.CUPS, s.FechaActivacion, s.RazonSocial, s.Identificador, s.informado_alta,"
                    + " p.DesProvincia as Provincia, s.DescripcionPoblacion, ec.es_incidencia";

                for (int i = 1; i < 7; i++)
                    strSql += " ,s.PotenciaPeriodo" + i;

                strSql += " from cont.eexxi_solicitudes_tmp s inner join"
                    + " cont.eexxi_param_solicitudes_codigos sc on"
                    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                    + " sc.CodigoDePaso = s.CodigoDePaso"
                    + " left outer join eexxi_param_provincias p on"
                    + " s.Provincia = p.CodigoPostal"
                    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                    //+ " h.cups = substr(s.CUPS, 1, 20)"
                    + " h.cups = s.CUPS"
                    // 20221129 ADD
                    + " and h.codigo_solicitud = s.CodigoDeSolicitud"
                    // 20221129 END ADD
                    + " LEFT OUTER JOIN eexxi_param_estados_casos ec ON"
                    + " ec.estado_id = h.estado_id"
                    + " LEFT OUTER JOIN cont.eexxi_solicitudes_t101 t101 ON"
                    + " h.cups = t101.CUPS AND"
                    + " h.codigo_solicitud = t101.CodigoDeSolicitud AND"
                    + " h.fecha_alta = t101.FechaActivacion"
                    + " where sc.Descripcion = 'ALTA'"
                    //+ " AND s.informado_alta is null"
                    + " order by s.FechaActivacion desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Inventario_Tabla c = new EndesaEntity.contratacion.Inventario_Tabla();
                    c.temporal = true;
                    c.cups22 = r["CUPS"].ToString();

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["nombre_cliente"] != System.DBNull.Value)
                        c.razon_social = r["nombre_cliente"].ToString();

                    if (r["es_incidencia"] != System.DBNull.Value)
                        c.tiene_incidencia = r["es_incidencia"].ToString() == "S";
                    
                    if (r["FechaActivacion"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["FechaActivacion"]);

                    if (r["RazonSocial"] != System.DBNull.Value)
                        if(c.razon_social == null)
                            c.razon_social = r["RazonSocial"].ToString();

                    if (r["informado_alta"] != System.DBNull.Value)
                        c.informado_alta = Convert.ToDateTime(r["informado_alta"]);

                    if (r["Identificador"] != System.DBNull.Value)
                        if(c.nif == null)                        
                            c.nif = r["Identificador"].ToString();

                    if (r["Provincia"] != System.DBNull.Value)
                        c.cod_provincia = r["Provincia"].ToString();

                    if (r["DescripcionPoblacion"] != System.DBNull.Value)
                        c.descripcion_poblacion = r["DescripcionPoblacion"].ToString();

                    for (int i = 1; i < 7; i++)
                        if (r["PotenciaPeriodo" + i] != System.DBNull.Value)
                            c.potencias[i - 1] = Convert.ToDouble(r["PotenciaPeriodo" + i]);

                    EndesaEntity.contratacion.Inventario_Tabla o;
                    //if (!dic_tmp_altas.TryGetValue(c.cups22.Substring(0, 20), out o))
                    if (!dic_tmp_altas.TryGetValue(c.cups22, out o))
                        //dic_tmp_altas.Add(c.cups22.Substring(0, 20), c);
                        dic_tmp_altas.Add(c.cups22, c);
                }
                db.CloseConnection();

                #endregion

                #region ALTAS NO TEMP
                strSql = "select t101.Identificador as nif, t101.RazonSocial as nombre_cliente,"
                    + " s.CUPS, s.FechaActivacion, s.RazonSocial, s.Identificador, s.informado_alta,"
                    + " p.DesProvincia as Provincia, s.DescripcionPoblacion, ec.es_incidencia";

                for (int i = 1; i < 7; i++)
                    strSql += " ,s.PotenciaPeriodo" + i;

                strSql += " from cont.eexxi_solicitudes s inner join"
                    + " cont.eexxi_param_solicitudes_codigos sc on"
                    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                    + " sc.CodigoDePaso = s.CodigoDePaso"
                    + " left outer join eexxi_param_provincias p on"
                    + " s.Provincia = p.CodigoPostal"
                    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                    //+ " h.cups = substr(s.CUPS, 1, 20)"                    
                    + " h.cups = s.CUPS"
                    // 20221129 ADD
                    + " and h.codigo_solicitud = s.CodigoDeSolicitud"
                    // 20221129 END ADD
                    + " LEFT OUTER JOIN eexxi_param_estados_casos ec ON"
                    + " ec.estado_id = h.estado_id"
                    + " LEFT OUTER JOIN cont.eexxi_solicitudes_t101 t101 ON"
                    + " h.cups = t101.CUPS AND"
                    + " h.codigo_solicitud = t101.CodigoDeSolicitud AND"
                    + " h.fecha_alta = t101.FechaActivacion"
                    + " where sc.Descripcion = 'ALTA' AND"
                    + " s.informado_alta is null"
                    + " order by s.FechaActivacion desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Inventario_Tabla c = new EndesaEntity.contratacion.Inventario_Tabla();
                    c.temporal = true;

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["nombre_cliente"] != System.DBNull.Value)
                        c.razon_social = r["nombre_cliente"].ToString();

                    if (r["es_incidencia"] != System.DBNull.Value)
                        c.tiene_incidencia = r["es_incidencia"].ToString() == "S";                    

                    c.cups22 = r["CUPS"].ToString();
                    if (r["FechaActivacion"] != System.DBNull.Value)
                        c.fecha_alta = Convert.ToDateTime(r["FechaActivacion"]);

                    if (r["informado_alta"] != System.DBNull.Value)
                        c.informado_alta = Convert.ToDateTime(r["informado_alta"]);

                    if (r["RazonSocial"] != System.DBNull.Value)
                        if(c.razon_social == null)
                            c.razon_social = r["RazonSocial"].ToString();

                    if (r["Identificador"] != System.DBNull.Value)
                        if(c.nif == null)                        
                            c.nif = r["Identificador"].ToString();

                    if (r["Provincia"] != System.DBNull.Value)
                        c.cod_provincia = r["Provincia"].ToString();

                    if (r["DescripcionPoblacion"] != System.DBNull.Value)
                        c.descripcion_poblacion = r["DescripcionPoblacion"].ToString();

                    for (int i = 1; i < 7; i++)
                        if (r["PotenciaPeriodo" + i] != System.DBNull.Value)
                            c.potencias[i - 1] = Convert.ToDouble(r["PotenciaPeriodo" + i]);

                    EndesaEntity.contratacion.Inventario_Tabla o;
                    //if (!dic_tmp_altas.TryGetValue(c.cups22.Substring(0, 20), out o))
                    if (!dic_tmp_altas.TryGetValue(c.cups22, out o))
                        //dic_tmp_altas.Add(c.cups22.Substring(0, 20), c);
                        dic_tmp_altas.Add(c.cups22, c);
                }
                db.CloseConnection();
                #endregion

                #region BAJAS TEMP
                strSql = "select t101.Identificador as nif, t101.RazonSocial as nombre_cliente,"
                    + " s.CUPS, s.FechaActivacion, ec.es_incidencia, s.informado_alta"
                    + " from cont.eexxi_solicitudes_tmp s inner join"
                    + " cont.eexxi_param_solicitudes_codigos sc on"
                    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                    + " sc.CodigoDePaso = s.CodigoDePaso"
                    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                    // + " h.cups = substr(s.CUPS, 1, 20)"
                    + " h.cups = s.CUPS"
                    // 20221129 ADD
                    + " and h.codigo_solicitud = s.CodigoDeSolicitud"
                    // 20221129 END ADD
                    + " INNER JOIN eexxi_param_estados_casos ec ON"
                    + " ec.estado_id = h.estado_id"
                    + " LEFT OUTER JOIN cont.eexxi_solicitudes_t101 t101 ON"
                    + " h.cups = t101.CUPS AND"
                    + " h.codigo_solicitud = t101.CodigoDeSolicitud AND"
                    + " h.fecha_alta = t101.FechaActivacion"
                    + " where sc.Descripcion = 'BAJA' AND"
                    + " s.informado_alta is null"
                    + " order by s.FechaActivacion desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Inventario_Tabla c = new EndesaEntity.contratacion.Inventario_Tabla();
                    c.temporal = true;

                    if (r["nif"] != System.DBNull.Value)
                        c.nif = r["nif"].ToString();

                    if (r["nombre_cliente"] != System.DBNull.Value)
                        c.razon_social = r["nombre_cliente"].ToString();

                    if (r["es_incidencia"] != System.DBNull.Value)
                        c.tiene_incidencia = r["es_incidencia"].ToString() == "S";

                    c.cups22 = r["CUPS"].ToString();

                    if (r["FechaActivacion"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["FechaActivacion"]);

                    if (r["informado_alta"] != System.DBNull.Value)
                        c.informado_alta = Convert.ToDateTime(r["informado_alta"]);

                    EndesaEntity.contratacion.Inventario_Tabla o;
                    //if (!dic_tmp_bajas.TryGetValue(c.cups22.Substring(0, 20), out o))
                    if (!dic_tmp_bajas.TryGetValue(c.cups22, out o))
                        //dic_tmp_bajas.Add(c.cups22.Substring(0, 20), c);
                        dic_tmp_bajas.Add(c.cups22, c);
                    else                    
                        o.temporal = true;
                    
                }
                db.CloseConnection();

                #endregion

                #region ALTAS NO TEMP
                strSql = "select s.CUPS, s.FechaActivacion, ec.es_incidencia, s.informado_alta"
                    + " from cont.eexxi_solicitudes s inner join"
                    + " cont.eexxi_param_solicitudes_codigos sc on"
                    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                    + " sc.CodigoDePaso = s.CodigoDePaso"
                    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                    //+ " h.cups = substr(s.CUPS, 1, 20)"
                    + " h.cups = s.CUPS"
                    // 20221129 ADD
                    + " and h.codigo_solicitud = s.CodigoDeSolicitud"
                    // 20221129 END ADD
                    + " INNER JOIN eexxi_param_estados_casos ec ON"
                    + " ec.estado_id = h.estado_id"
                    + " where sc.Descripcion = 'BAJA' AND"
                    + " s.informado_alta is null"
                    + " order by s.FechaActivacion desc";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Inventario_Tabla c = new EndesaEntity.contratacion.Inventario_Tabla();
                    c.temporal = true;

                    if (r["es_incidencia"] != System.DBNull.Value)
                        c.tiene_incidencia = r["es_incidencia"].ToString() == "S";

                    c.cups22 = r["CUPS"].ToString();

                    if (r["FechaActivacion"] != System.DBNull.Value)
                        c.fecha_baja = Convert.ToDateTime(r["FechaActivacion"]);

                    if (r["informado_alta"] != System.DBNull.Value)
                        c.informado_alta = Convert.ToDateTime(r["informado_alta"]);

                    EndesaEntity.contratacion.Inventario_Tabla o;
                    //if (!dic_tmp_bajas.TryGetValue(c.cups22.Substring(0, 20), out o))
                    if (!dic_tmp_bajas.TryGetValue(c.cups22, out o))
                        //dic_tmp_bajas.Add(c.cups22.Substring(0, 20), c);
                        dic_tmp_bajas.Add(c.cups22, c);
                    else
                        o.temporal = true;


                }
                db.CloseConnection();
                #endregion

                foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario_Tabla> p in dic_tmp_bajas)
                {
                    EndesaEntity.contratacion.Inventario_Tabla o;
                    //if (dic_tmp_altas.TryGetValue(p.Key.Substring(0, 20), out o))
                    if (dic_tmp_altas.TryGetValue(p.Key, out o))
                    {
                        o.fecha_baja = o.fecha_baja;
                        o.vigente = false;
                        o.temporal = true;
                        o.informado_alta = p.Value.informado_alta;
                    }
                    else
                    {
                        //if (dic.TryGetValue(p.Key.Substring(0, 20), out o))
                        if (dic.TryGetValue(p.Key, out o))
                        {
                            o.fecha_baja = p.Value.fecha_baja;
                            o.temporal = true; 
                            o.vigente = false;
                            o.informado_alta = p.Value.informado_alta;
                            //dic_tmp_altas.Add(p.Key.Substring(0, 20), o);
                            dic_tmp_altas.Add(p.Key, o);
                        }
                    }

                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "EEXXI - CargaInventario",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public void GetInfoCUPS20(string cups22)
        {
            EndesaEntity.contratacion.Inventario_Tabla o;
            if (dic_tmp_altas.TryGetValue(cups22, out o))
            {

                this.nif = o.nif;
                this.razon_social = o.razon_social;
                this.fecha_baja = o.fecha_baja;
            }
            else
            {
                if (dic.TryGetValue(cups22, out o))
                {
                    this.nif = o.nif;
                    this.razon_social = o.razon_social;
                    this.fecha_baja = o.fecha_baja;
                }
            }
        }

        public void Inventario_a_Excel(string fichero)
        {
            int f = 0;
            int c = 0;

            FileInfo fileInfo = new FileInfo(fichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("Puntos Vigentes XXI");

            var headerCells = workSheet.Cells[1, 1, 1, 4];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
            workSheet.Cells[f, c].Value = "CUPS"; c++;
            workSheet.Cells[f, c].Value = "FECHA ALTA"; c++;


            foreach (KeyValuePair<string, EndesaEntity.contratacion.Inventario_Tabla> p in dic)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.Value.nif; c++;
                workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
                workSheet.Cells[f, c].Value = p.Value.cups22; c++;

                if (p.Value.fecha_alta > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.Value.fecha_alta;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;


            }

            var allCells = workSheet.Cells[1, 1, f, 10];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:D1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();
        }

        public void Update_Informado_Alta()
        {
            MySQLDB db;
            MySqlCommand command;            
            string strSql;

            try
            {
                strSql = "update eexxi_solicitudes_tmp set "
                    + " informado_alta = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                    + " where informado_alta is null";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "update eexxi_solicitudes set "
                    + " informado_alta = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                    + " where informado_alta is null";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();


                //strSql = "update eexxi_solicitudes_tmp s inner join"
                //    + " cont.eexxi_param_solicitudes_codigos sc on"
                //    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                //    + " sc.CodigoDePaso = s.CodigoDePaso"
                //    + " left outer join eexxi_param_provincias p on"
                //    + " s.Provincia = p.CodigoPostal"
                //    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                //    + " h.cups = substr(s.CUPS, 1, 20)"
                //    + " LEFT OUTER JOIN eexxi_param_estados_casos ec ON"
                //    + " ec.estado_id = h.estado_id"
                //    + " set s.informado_alta = '"
                //    + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                //    + " where sc.Descripcion = 'ALTA' AND"
                //    + " s.informado_alta is null";
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "update eexxi_solicitudes s inner join"
                //    + " cont.eexxi_param_solicitudes_codigos sc on"
                //    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                //    + " sc.CodigoDePaso = s.CodigoDePaso"
                //    + " left outer join eexxi_param_provincias p on"
                //    + " s.Provincia = p.CodigoPostal"
                //    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                //    + " h.cups = substr(s.CUPS, 1, 20)"
                //    + " LEFT OUTER JOIN eexxi_param_estados_casos ec ON"
                //    + " ec.estado_id = h.estado_id"
                //    + " set s.informado_alta = '"
                //    + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                //    + " where sc.Descripcion = 'ALTA' AND"
                //    + " s.informado_alta is null";                    
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "update eexxi_solicitudes_tmp s inner join"
                //    + " cont.eexxi_param_solicitudes_codigos sc on"
                //    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                //    + " sc.CodigoDePaso = s.CodigoDePaso"
                //    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                //    + " h.cups = substr(s.CUPS, 1, 20)"
                //    + " LEFT OUTER JOIN eexxi_param_estados_casos ec ON"
                //    + " ec.estado_id = h.estado_id"
                //    + " set s.informado_alta = '"
                //    + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                //    + " where sc.Descripcion = 'BAJA'"
                //    + " order by s.FechaActivacion desc";
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();

                //strSql = "update eexxi_solicitudes s inner join"
                //    + " cont.eexxi_param_solicitudes_codigos sc on"
                //    + " sc.CodigoDelProceso = s.CodigoDelProceso and"
                //    + " sc.CodigoDePaso = s.CodigoDePaso"
                //    + " LEFT OUTER JOIN eexxi_casos_hist h ON"
                //    + " h.cups = substr(s.CUPS, 1, 20)"
                //    + " LEFT OUTER JOIN eexxi_param_estados_casos ec ON"
                //    + " ec.estado_id = h.estado_id"
                //    + " set s.informado_alta = '"
                //    + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                //    + " where sc.Descripcion = 'BAJA'"
                //    + " order by s.FechaActivacion desc";
                //db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                //command = new MySqlCommand(strSql, db.con);
                //command.ExecuteNonQuery();
                //db.CloseConnection();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                "EEXXI - Inventario - Update_Informado_Alta",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

    }
}

