using EndesaBusiness.servidores;
using Google.Protobuf.WellKnownTypes;
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
    public class Casos : EndesaEntity.contratacion.xxi.Casos_Tabla
    {
        public Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> dic { get; set; }
        public Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> dic_casos_altas { get; set; }
        List<EndesaEntity.contratacion.xxi.Casos_Tabla> lista;
        Dictionary<string, string> dic_ps_at_hist;

        contratacion.PS_AT ps_at;

        EndesaBusiness.contratacion.eexxi.Comentarios comments;

        public Casos()
        {
            lista = Carga();
            dic_casos_altas = CargaDicAltasCasos();
            comments = new Comentarios();
        }

        public Casos(contratacion.eexxi.Inventario inventario)
        {
            
            ps_at = new PS_AT();
            comments = new Comentarios();

            // Cargamos la lista de casos abiertos
            dic = DicCasosAbiertos();
            // Completamos datos con PS_HIST
            contratacion.PS_AT_HIST ps_at_hist = new PS_AT_HIST(dic.Values.Select(z => z.cups).ToList());

            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> p in dic)
            {               

                List<EndesaEntity.contratacion.PS_AT_Tabla> lista;
                if(ps_at_hist.dic.TryGetValue(p.Key, out lista))
                    foreach(EndesaEntity.contratacion.PS_AT_Tabla pp in lista)
                    {
                        if (pp.empresa == "EEXXI" && pp.fecha_alta_contrato == p.Value.fecha_alta)
                        {
                            p.Value.existe_alta = true;                
                            break;
                        }
                    }
                                 
                                  
                

                if (ps_at.ExisteAlta(p.Value.cups))
                {
                    if(p.Value.nif == null || p.Value.nif == "")
                        p.Value.nif = ps_at.cif;


                    if (p.Value.razon_social == null || p.Value.razon_social == "")
                        p.Value.razon_social = ps_at.nombre_cliente;


                    p.Value.empresa_ps = ps_at.empresa;
                    p.Value.estado_contrato_ps = ps_at.estado_contrato_descripcion;
                }
                else
                {
                    inventario.GetInfoCUPS20(p.Value.cups);
                    if(p.Value.nif == null || p.Value.nif == "")
                    {
                        p.Value.nif = inventario.nif;
                        p.Value.fecha_baja = inventario.fecha_baja;
                    }
                        

                    if (p.Value.razon_social == null || p.Value.razon_social == "")
                        p.Value.razon_social = inventario.razon_social;
                }

            }
                        
            dic_casos_altas = CargaDicAltasCasos();
            
        }

        public Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> CargaDicAltasCasos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> dic = 
                new Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico>();
                        
            strSql = "SELECT s.CUPS, s.CodigoDeSolicitud, h.documentacion_impresa"
                + " FROM eexxi_solicitudes s"
                + " INNER JOIN eexxi_param_solicitudes_codigos c ON"
                + " c.CodigoDelProceso = s.CodigoDelProceso AND"
                + " c.CodigoDePaso = s.CodigoDePaso"
                + " LEFT OUTER JOIN"
                + " eexxi_casos_hist h ON"                
                + " h.cups = s.CUPS AND"
                + " h.codigo_solicitud = s.CodigoDeSolicitud"
                + " WHERE c.Descripcion = 'ALTA'"
                + " ORDER BY s.Identificador, s.last_update_date";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico c = new EndesaEntity.contratacion.xxi.Casos_Tabla_Historico();
                c.cups = r["CUPS"].ToString();
                c.codigo_solicitud = r["CodigoDeSolicitud"].ToString();
                c.documentacion_impresa = r["documentacion_impresa"].ToString() == "S";

                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico o;
                if (!dic.TryGetValue(c.cups, out o))
                {
                    dic.Add(c.cups, c);
                }
                else
                {
                    o.codigo_solicitud = c.codigo_solicitud;
                    o.documentacion_impresa = c.documentacion_impresa;
                }


            }
            db.CloseConnection();

            
            strSql = "SELECT s.CUPS, s.CodigoDeSolicitud, h.documentacion_impresa"
                + " FROM eexxi_solicitudes_tmp s"
                + " INNER JOIN eexxi_param_solicitudes_codigos c ON"
                + " c.CodigoDelProceso = s.CodigoDelProceso AND"
                + " c.CodigoDePaso = s.CodigoDePaso"
                + " LEFT OUTER JOIN"
                + " eexxi_casos_hist h ON"                
                + " h.cups = s.CUPS AND"
                + " h.codigo_solicitud = s.CodigoDeSolicitud"
                + " WHERE c.Descripcion = 'ALTA'"
                + " ORDER BY s.Identificador, s.last_update_date";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico c = new EndesaEntity.contratacion.xxi.Casos_Tabla_Historico();
                c.cups = r["CUPS"].ToString();
                c.codigo_solicitud = r["CodigoDeSolicitud"].ToString();
                c.documentacion_impresa = r["documentacion_impresa"].ToString() == "S";

                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico o;
                if (!dic.TryGetValue(c.cups, out o))
                {
                    dic.Add(c.cups, c);
                }
                else
                {
                    o.codigo_solicitud = c.codigo_solicitud;
                    o.documentacion_impresa = c.documentacion_impresa;                    
                }


            }
            db.CloseConnection();

            return dic;




        }

        public void GetCaso(string tipo, bool existe_baja, bool existe_ps, string empresa_ps, int estado_contrato_ps_id)
        {
            EndesaEntity.contratacion.xxi.Casos_Tabla e = lista.FirstOrDefault(z => z.tipo == tipo && z.existe_baja == existe_baja && z.existe_ps == existe_ps
                && z.empresa_ps == empresa_ps && z.estado_contrato_ps_id == estado_contrato_ps_id);

            if (e != null)
            {
                this.caso = e.caso;
                this.existe_baja = e.existe_baja;
                this.existe_ps = e.existe_ps;
                this.crear_contrato = e.crear_contrato;
                this.crear_incidencia = e.crear_incidencia;
                this.realizar_seguimiento = e.realizar_seguimiento;
                this.acciones = e.acciones;
                this.observaciones = e.observaciones;
            }
            else
            {
                this.caso = 99;
                this.crear_contrato = false;
                this.crear_incidencia = true;
                this.realizar_seguimiento = true;
                this.existe_baja = existe_baja;
                this.existe_ps = existe_ps;
                this.acciones = "CASO NO CONTEMPLADO";
                this.observaciones = "CASO NO CONTEMPLADO";
            }
        }

        private List<EndesaEntity.contratacion.xxi.Casos_Tabla> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            List<EndesaEntity.contratacion.xxi.Casos_Tabla> lista = new List<EndesaEntity.contratacion.xxi.Casos_Tabla>();

            try
            {

                strSql = "select caso, tipo, existe_baja, existe_ps, empresa_ps,"
                    + " estado_contrato_ps_id, estado_contrato_ps_descripcion,"
                    + " crear_contrato, crear_incidencia, realizar_seguimiento,"
                    + " acciones, observaciones"
                    + " from eexxi_casos";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Casos_Tabla c = new EndesaEntity.contratacion.xxi.Casos_Tabla();

                    c.caso = Convert.ToInt32(r["caso"]);
                    c.tipo = r["tipo"].ToString();
                    c.existe_baja = r["existe_baja"].ToString() == "S";
                    c.existe_ps = r["existe_ps"].ToString() == "S";

                    if (r["empresa_ps"] != System.DBNull.Value)
                        if (r["empresa_ps"].ToString() != "")
                            c.empresa_ps = r["empresa_ps"].ToString();

                    c.estado_contrato_ps_id = Convert.ToInt32(r["estado_contrato_ps_id"]);

                    if (r["estado_contrato_ps_descripcion"] != System.DBNull.Value)
                        c.estado_contrato_ps_descripcion = r["estado_contrato_ps_descripcion"].ToString();

                    c.crear_contrato = r["crear_contrato"].ToString() == "S";
                    c.crear_incidencia = r["crear_incidencia"].ToString() == "S";
                    c.realizar_seguimiento = r["realizar_seguimiento"].ToString() == "S";

                    if (r["acciones"] != System.DBNull.Value)
                        c.acciones = r["acciones"].ToString();

                    if (r["observaciones"] != System.DBNull.Value)
                        c.observaciones = r["observaciones"].ToString();

                    lista.Add(c);

                }

                db.CloseConnection();
                return lista;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Casis - Carga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                return null;
            }
        }

        private Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> DicCasosAbiertos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> dic = 
                new Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico>();
            try
            {

                // Cargamos los comentarios
                //Comentarios coments = new Comentarios(true);

                strSql = "select t.Identificador, t.RazonSocial,"
                    + " h.cups, h.codigo_solicitud, h.fecha_alta, h.fecha_baja, h.existe_alta, h.existe_ps,"                                
                    + " h.empresa_ps, h.estado_contrato_ps, h.crear_contrato, h.crear_incidencia,"
                    + " h.realizar_seguimiento, h.acciones, h.documentacion_impresa, h.estado_id, pc.descripcion"
                    + " FROM eexxi_casos_hist h "
                    + " LEFT OUTER JOIN cont.eexxi_solicitudes_t101 t on"                    
                    + " t.CUPS = h.cups AND"
                    + " t.CodigoDeSolicitud  = h.codigo_solicitud AND"
                    + " t.FechaActivacion = h.fecha_alta"
                    + " inner join eexxi_param_estados_casos pc on"
                    + " pc.estado_id = h.estado_id"                    
                    + " ORDER BY h.created_date desc";
                    
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Casos_Tabla_Historico c = 
                        new EndesaEntity.contratacion.xxi.Casos_Tabla_Historico();

                    if (r["CUPS"] != System.DBNull.Value)
                    {

                        if (r["Identificador"] != System.DBNull.Value)
                            c.nif = r["Identificador"].ToString();

                        if (r["RazonSocial"] != System.DBNull.Value)
                            c.razon_social = r["RazonSocial"].ToString();

                        c.cups = r["cups"].ToString();
                        c.codigo_solicitud = r["codigo_solicitud"].ToString();
                        c.fecha_alta = Convert.ToDateTime(r["fecha_alta"]);

                        if (r["fecha_baja"] != System.DBNull.Value)
                            c.fecha_baja = Convert.ToDateTime(r["fecha_baja"]);
                        else
                        {

                        }

                        c.crear_contrato = r["crear_contrato"].ToString() == "S";
                        c.crear_incidencia = r["crear_incidencia"].ToString() == "S";
                        c.realizar_seguimiento = r["realizar_seguimiento"].ToString() == "S";
                        c.documentacion_impresa = r["documentacion_impresa"].ToString() == "S";

                        if (r["acciones"] != System.DBNull.Value)
                            c.acciones = r["acciones"].ToString();

                        c.estado_id = Convert.ToInt32(r["estado_id"]);
                        c.descripcion_estado_caso = r["descripcion"].ToString();

                        c.comentario = comments.UltimoComentario(c.cups.Substring(0, 20), c.codigo_solicitud);

                        EndesaEntity.contratacion.xxi.Casos_Tabla_Historico o;
                        if (!dic.TryGetValue(c.cups, out o))
                        {
                            dic.Add(c.cups, c);
                        }
                        else
                        {
                            if(c.fecha_alta == o.fecha_alta && 
                                (!o.documentacion_impresa  && c.documentacion_impresa))
                            {
                                o.crear_contrato = c.crear_contrato;
                                o.crear_incidencia = c.crear_incidencia;
                                o.realizar_seguimiento = c.realizar_seguimiento;
                                o.documentacion_impresa = c.documentacion_impresa;

                            }
                        }
                    }

                    



                }
                db.CloseConnection();

                Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> dic2 =
                    new Dictionary<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico>();
                foreach(KeyValuePair<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> d in dic)
                {
                    if (!d.Value.documentacion_impresa || d.Value.estado_id != 2)
                        dic2.Add(d.Key, d.Value);
                }


                return dic2;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "Contratos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return null;
            }
        }

        public void ActualizaEstadoCaso(string cups, string codigo_solicitud, int estado_id)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

             strSql = "update eexxi_casos_hist set estado_id = " + estado_id + ","
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                + " where cups = '" + cups + "' and"
                + " codigo_solicitud = '" + codigo_solicitud + "'";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void MarcaComoImpreso(string cups, string codigo_solicitud)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update eexxi_casos_hist set documentacion_impresa = 'S',"
                + "last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                //+ " where cups = '" + cups.Substring(0, 20) + "' and"
                + " where cups = '" + cups + "' and"
                + " codigo_solicitud = '" + codigo_solicitud + "'";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void MarcaComoImpreso(string cups)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update eexxi_casos_hist set documentacion_impresa = 'S',"
                + "last_update_by = '" + System.Environment.UserName.ToUpper() + "'"
                //+ " where cups = '" + cups.Substring(0, 20) + "' and"
                + " where cups = '" + cups + "' and"
                + " documentacion_impresa = 'N'";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void GuardaCasos(List<EndesaEntity.contratacion.xxi.InformeCasos> lista)
        {

            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int numReg = 0;

            for (int i = 0; i < lista.Count; i++)
            {
                numReg++;
                if (firstOnly)
                {

                    sb.Append("replace into eexxi_casos_hist (cups, codigo_solicitud, fecha_alta, fecha_baja, existe_alta,");
                    sb.Append("existe_ps, empresa_ps, estado_contrato_ps, crear_contrato, crear_incidencia,");
                    sb.Append("realizar_seguimiento, acciones, documentacion_impresa, ");
                    sb.Append("created_by, created_date) values");
                    firstOnly = false;
                }



                //sb.Append("('").Append(lista[i].cups20).Append("',");
                sb.Append("('").Append(lista[i].cups22).Append("',");

                if (lista[i].codigo_solicitud != null)
                    sb.Append("'").Append(lista[i].codigo_solicitud).Append("',");
                else
                    sb.Append("null,");

                if (lista[i].fecha_activacion > DateTime.MinValue)
                    sb.Append("'").Append(lista[i].fecha_activacion.ToString("yyyy-MM-dd")).Append("',");
                else
                    sb.Append("null,");

                if (lista[i].fecha_baja > DateTime.MinValue)
                    sb.Append("'").Append(lista[i].fecha_baja.ToString("yyyy-MM-dd")).Append("',");
                else
                    sb.Append("null,");

                sb.Append("'").Append(lista[i].existe_ps_hist ? "S" : "N").Append("',");
                sb.Append("'").Append(lista[i].existe_ps ? "S" : "N").Append("',");

                if (lista[i].empresa != null)
                    sb.Append("'").Append(lista[i].empresa).Append("',");
                else
                    sb.Append("null,");

                if (lista[i].estado_contrato_ps != null)
                    sb.Append("'").Append(lista[i].estado_contrato_ps).Append("',");
                else
                    sb.Append("null,");

                sb.Append("'").Append(lista[i].crear_contrato ? "S" : "N").Append("',");
                sb.Append("'").Append(lista[i].crear_incidencia ? "S" : "N").Append("',");
                sb.Append("'").Append(lista[i].realizar_seguimiento ? "S" : "N").Append("',");

                if (lista[i].acciones != null)
                    sb.Append("'").Append(lista[i].acciones).Append("',");
                else
                    sb.Append("null,");

                EndesaEntity.contratacion.xxi.Casos_Tabla_Historico o;
                if (dic_casos_altas.TryGetValue(lista[i].cups20, out o))
                    sb.Append("'").Append(o.documentacion_impresa ? "S" : "N").Append("',");
                else
                    sb.Append("'N',");

                sb.Append("'").Append(System.Environment.UserName).Append("','").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("'),");

                if (numReg == 250)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    numReg = 0;
                }

            }

            if (numReg > 0)
            {
                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                sb = null;
                numReg = 0;
            }
        }

        private void ActualizaCasos()
        {

        }


        public void CasosAbiertos_a_Excel(string fichero)
        {
            int f = 0;
            int c = 0;


            FileInfo fileInfo = new FileInfo(fichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CASOS");

            var headerCells = workSheet.Cells[1, 1, 1, 30];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "NIF"; c++;
            workSheet.Cells[f, c].Value = "RAZÓN SOCIAL"; c++;
            workSheet.Cells[f, c].Value = "CUPS"; c++;
            workSheet.Cells[f, c].Value = "FECHA ALTA"; c++;
            workSheet.Cells[f, c].Value = "FECHA BAJA"; c++;
            workSheet.Cells[f, c].Value = "EXISTE ALTA EEXXI"; c++;
            workSheet.Cells[f, c].Value = "EMPRESA PS"; c++;
            workSheet.Cells[f, c].Value = "ESTADO CONTRATO PS"; c++;
            workSheet.Cells[f, c].Value = "DOCUMENTACIÓN IMPRESA"; c++;
            workSheet.Cells[f, c].Value = "ESTADO"; c++;

            foreach (KeyValuePair<string, EndesaEntity.contratacion.xxi.Casos_Tabla_Historico> p in dic)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.Value.nif; c++;
                workSheet.Cells[f, c].Value = p.Value.razon_social; c++;
                workSheet.Cells[f, c].Value = p.Value.cups; c++;

                if (p.Value.fecha_alta > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.Value.fecha_alta;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;
                if (p.Value.fecha_baja > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.Value.fecha_baja;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                workSheet.Cells[f, c].Value = p.Value.existe_alta; c++;
                workSheet.Cells[f, c].Value = p.Value.empresa_ps; c++;
                workSheet.Cells[f, c].Value = p.Value.estado_contrato_ps; c++;
                workSheet.Cells[f, c].Value = p.Value.documentacion_impresa; c++;
                workSheet.Cells[f, c].Value = p.Value.descripcion_estado_caso; c++;

            }

            var allCells = workSheet.Cells[1, 1, f, 10];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:J1"].AutoFilter = true;
            allCells.AutoFitColumns();
            excelPackage.Save();


        }
    }
}
