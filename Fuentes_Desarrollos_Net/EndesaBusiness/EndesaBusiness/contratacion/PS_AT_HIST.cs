using EndesaBusiness.servidores;
using EndesaEntity;
using MySql.Data.MySqlClient;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class PS_AT_HIST : EndesaEntity.contratacion.PS_AT_Tabla
    {
        //public Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla> dic { get; set; }
        public Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>> dic { get; set; }


        public PS_AT_HIST()
        {

        }

        public PS_AT_HIST(List<string> lista_cups_20)
        {

            //dic = new Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla>();
            dic = new Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>>();
            if (lista_cups_20.Count > 0)
                Carga(lista_cups_20, null);
        }

        public PS_AT_HIST(List<string> lista_cups_20, string empresa)
        {

            //dic = new Dictionary<string, EndesaEntity.contratacion.PS_AT_Tabla>();
            dic = new Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>>();
            if (lista_cups_20.Count > 0)
                Carga(lista_cups_20, empresa);
        }

        public List<EndesaEntity.contratacion.PS_AT_Tabla> PS_AT_HIST_CUPS(string cups)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            List<EndesaEntity.contratacion.PS_AT_Tabla> lista = new List<EndesaEntity.contratacion.PS_AT_Tabla>();

            strSql = "SELECT h.EMPRESA, h.CUPS22, h.CCONTATR, MIN(h.Fecha_Anexion) AS min_fecha_anexion,"
                + " MAX(h.Fecha_Anexion) AS max_fecha_anexion"
                + " FROM cont.PS_AT_HIST h WHERE"
                + " h.CUPS22 = '" + cups + "' GROUP BY h.EMPRESA ORDER BY h.Fecha_Anexion";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.contratacion.PS_AT_Tabla c = new EndesaEntity.contratacion.PS_AT_Tabla();
                c.empresa = r["EMPRESA"].ToString();
                c.cups22 = r["CUPS22"].ToString();
                c.contrato_atr = Convert.ToInt64(r["CCONTATR"]);
                c.min_fecha_anexion = Convert.ToDateTime(r["min_fecha_anexion"]);
                c.max_fecha_anexion = Convert.ToDateTime(r["max_fecha_anexion"]);

                lista.Add(c);

            }
            db.CloseConnection();

            return lista;
        }

        private void Carga(List<string> lista_cups_20, string empresa)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            
            strSql = "SELECT cups20, CUPS22, cliente, NIF, DATE_FORMAT(fAltaCont, '%Y-%m-%d') AS fecha_alta, EMPRESA,"
                + " Fecha_Anexion"
                + " from PS_AT_HIST"
                + " where cups20 in ('" + lista_cups_20[0].Substring(0,20) + "'";

            for (int i = 1; i < lista_cups_20.Count; i++)
                strSql += " ,'" + lista_cups_20[i].Substring(0, 20) + "'";

            strSql += ") and (estadoCont = '003' || estadoCont = '004')";
            if (empresa == null)
                strSql += " and EMPRESA = 'EEXXI'";
            else
                strSql += " and EMPRESA = '" + empresa + "'";

            // strSql += " GROUP BY cups20 order by fAltaCont desc";
            // Nos quedamos con el último contrato en vigor
            strSql += " order by cups20 , fAltaCont desc";
            

            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.contratacion.PS_AT_Tabla c = new EndesaEntity.contratacion.PS_AT_Tabla();
                if (r["fecha_alta"] != System.DBNull.Value)
                    c.fecha_alta_contrato = Convert.ToDateTime(r["fecha_alta"]);


                c.empresa = r["EMPRESA"].ToString();
                c.cups20 = r["cups20"].ToString();
                c.cups22 = r["CUPS22"].ToString();
                c.nombre_cliente = r["cliente"].ToString();
                c.cif = r["NIF"].ToString();

                if (r["Fecha_Anexion"] != System.DBNull.Value)
                    c.fecha_anexion = Convert.ToDateTime(r["Fecha_Anexion"]);


                //EndesaEntity.contratacion.PS_AT_Tabla o;
                List<EndesaEntity.contratacion.PS_AT_Tabla> o;
                if (!dic.TryGetValue(c.cups22, out o))
                {
                    o = new List<EndesaEntity.contratacion.PS_AT_Tabla>();
                    o.Add(c);
                    dic.Add(c.cups22, o);
                }
                else
                    o.Add(c);
                    
            }
            db.CloseConnection();
        }

        public bool ExisteAlta(string cups22)
        {
            bool existe = false;
            List<EndesaEntity.contratacion.PS_AT_Tabla> o;
            if (dic.TryGetValue(cups22, out o))
            {
                existe = true;
                this.empresa = o[0].empresa;
                this.fecha_alta_contrato = o[0].fecha_alta_contrato;
                this.cups20 = o[0].cups20;
                this.cups22 = o[0].cups22;
                this.nombre_cliente = o[0].nombre_cliente;
                this.cif = o[0].cif;
            }

            return existe;       
        
        }

        public List<DateTime> ListaFechaAnexion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            List<DateTime> l = new List<DateTime>();

            strSql = "SELECT h.Fecha_Anexion FROM PS_AT_HIST h GROUP BY h.Fecha_Anexion ORDER BY h.Fecha_Anexion DESC";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                l.Add(Convert.ToDateTime(r["Fecha_Anexion"]));
            }
            db.CloseConnection();

            return l;

        }

        public void Exporta_PS_AT_Excel(DateTime fecha, string archivo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;
            int f = 1;
            int c = 1;
            string cups13 = "";

            

            try
            {

                string ruta_salida_archivo = archivo;


                FileInfo fileInfo = new FileInfo(ruta_salida_archivo);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                var workSheet = excelPackage.Workbook.Worksheets.Add("PS_AT");
                var headerCells = workSheet.Cells[1, 1, 1, 25];
                var headerFont = headerCells.Style.Font;

                workSheet.View.FreezePanes(2, 1);

                headerFont.Bold = true;

                strSql = "SELECT ps.EMPRESA, ps.NIF, ps.Cliente,"
                    + " ps.IDU as CUPS13, ps.CUPS22, ps.DDISTRIB AS DISTRIBUIDORA,"
                    + " ps.CCONTATR AS CONTRATO_ATR, CNUMCATR AS NUM_CONTRATO_ATR,"
                    + " ps.fAltaCont AS F_ALTA_CONTRATO, fPrevBajaCont AS F_PREVISTA_BAJA,"
                    + " ps.fBajaCont AS F_BAJA_CONTRATO,"
                    + " ec.Descripcion AS ESTADO_CONTRATO,"
                    + " ps.FPREALTA, ps.TARIFA, ps.FPSERCON AS F_PUESTA_SERVICIO, ps.TENSION,"
                    + " ps.CONTREXT AS CONTRATO_EXTERNO,"
                    + " ps.Version AS VERSION_CONTRATO_EXTERNO,"
                    + " ps.VPOTCAL1 AS POTENCIA_CONTRATADA_P1,"
                    + " ps.VPOTCAL2 AS POTENCIA_CONTRATADA_P2,"
                    + " ps.VPOTCAL3 AS POTENCIA_CONTRATADA_P3,"
                    + " ps.VPOTCAL4 AS POTENCIA_CONTRATADA_P4,"
                    + " ps.VPOTCAL5 AS POTENCIA_CONTRATADA_P5,"
                    + " ps.VPOTCAL6 AS POTENCIA_CONTRATADA_P6,"
                    + " ps.TINDGCPY, tipo_contador.descripcion as TTICONPS, segmentoMercado,"
                    + " ps.municipio, ps.provincia,"
                    + " if (ps.tipoGestionATR = 1, 'No', 'Si') AS GESTION_PROPIA_ATR,"
                    + " ps.TPUNTMED AS TIPO_PUNTO_MEDIDA,"
                    + " ps.descripcion_autoconsumo, ps.f_ult_mod AS F_ULTIMA_ACTUALIZACION"
                    + " FROM cont.PS_AT_HIST AS ps"
                    + " LEFT OUTER JOIN cont.cont_estadoscontrato ec ON"
                    + " ec.Cod_Estado = ps.estadoCont"
                    + " LEFT OUTER JOIN cont.cont_ticonsps AS tipo_contador ON"
                    + " tipo_contador.tticonps = ps.TTICONPS"
                    + " WHERE ps.Fecha_Anexion = '" + fecha.ToString("yyyy-MM-dd") + "'"
                    + " ORDER BY NIF";


                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    c = 1;
                    #region Cabecera
                    if (firstOnly)
                    {
                        workSheet.Cells[f, c].Value = "EMPRESA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "NIF";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "RAZÓN SOCIAL";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CUPS13";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CUPS22";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TARIFA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "DISTRIBUIDORA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CONTRATO ATR";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "VERSIÓN CONTRATO ATR";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA ALTA CONTRATO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA PRE ALTA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA PUESTA EN SERVICIO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA PREVISTA BAJA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "FECHA BAJA CONTRATO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "ESTADO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TENSIÓN";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "CONTRATO EXTERNO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "VERSIÓN CONT EXTERNO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;



                        for (int i = 1; i <= 6; i++)
                        {
                            workSheet.Cells[f, c].Value = "POTENCIA CONTRATADA P" + i;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            c++;
                        }


                        workSheet.Cells[f, c].Value = "TINDGCPY";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TIPO CONTRATO PS";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "SEGMENTO MERCADO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "MUNICIPIO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "PROVINCIA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "GESTIÓN PROPIA ATR";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TIPO PUNTO MEDIDA";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "TIPO AUTOCONSUMO";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        workSheet.Cells[f, c].Value = "ÚLTIMA ACTUALIZACIÓN";
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray); c++;

                        firstOnly = false;

                    }
                    #endregion


                    c = 1;
                    f++;

                    if (r["EMPRESA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["EMPRESA"].ToString();
                    c++;

                    if (r["NIF"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["NIF"].ToString();
                    c++;

                    if (r["Cliente"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["Cliente"].ToString();


                    c++;

                    if (r["CUPS13"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["CUPS13"].ToString();
                        cups13 = r["CUPS13"].ToString();
                    }

                    c++;

                    if (r["CUPS22"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CUPS22"].ToString();
                    c++;

                    if (r["TARIFA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TARIFA"].ToString();
                    c++;

                    if (r["DISTRIBUIDORA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["DISTRIBUIDORA"].ToString();
                    c++;

                    if (r["CONTRATO_ATR"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CONTRATO_ATR"].ToString();
                    c++;

                    if (r["NUM_CONTRATO_ATR"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["NUM_CONTRATO_ATR"]);
                    c++;

                    if (r["F_ALTA_CONTRATO"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["F_ALTA_CONTRATO"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["F_ALTA_CONTRATO"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_ALTA_CONTRATO"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_ALTA_CONTRATO"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_ALTA_CONTRATO"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;


                    if (r["FPREALTA"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["FPREALTA"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["FPREALTA"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["FPREALTA"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["FPREALTA"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["FPREALTA"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;


                    if (r["F_PUESTA_SERVICIO"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["F_PUESTA_SERVICIO"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["F_PUESTA_SERVICIO"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_PUESTA_SERVICIO"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_PUESTA_SERVICIO"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_PUESTA_SERVICIO"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;



                    if (r["F_PREVISTA_BAJA"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["F_PREVISTA_BAJA"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        else if (Convert.ToInt32(r["F_PREVISTA_BAJA"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_PREVISTA_BAJA"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_PREVISTA_BAJA"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_PREVISTA_BAJA"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;

                    if (r["F_BAJA_CONTRATO"] != System.DBNull.Value)
                        if ((Convert.ToInt32(r["F_BAJA_CONTRATO"]) > 40000000))
                        {
                            fecha = new DateTime(4999, 12, 31);
                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        else if (Convert.ToInt32(r["F_BAJA_CONTRATO"]) > 0)
                        {
                            fecha = new DateTime(Convert.ToInt32(r["F_BAJA_CONTRATO"].ToString().Substring(0, 4)),
                                Convert.ToInt32(r["F_BAJA_CONTRATO"].ToString().Substring(4, 2)),
                                Convert.ToInt32(r["F_BAJA_CONTRATO"].ToString().Substring(6, 2)));

                            workSheet.Cells[f, c].Value = fecha;
                            workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    c++;


                    if (r["ESTADO_CONTRATO"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["ESTADO_CONTRATO"].ToString();
                    c++;

                    if (r["TENSION"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TENSION"].ToString();
                    c++;

                    if (r["CONTRATO_EXTERNO"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["CONTRATO_EXTERNO"].ToString();
                    c++;

                    if (r["VERSION_CONTRATO_EXTERNO"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["VERSION_CONTRATO_EXTERNO"]);
                    c++;

                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["POTENCIA_CONTRATADA_P" + i] != System.DBNull.Value)
                        {
                            if (Convert.ToInt32(r["POTENCIA_CONTRATADA_P" + i]) > 0)
                            {
                                workSheet.Cells[f, c].Value = Convert.ToDouble(r["POTENCIA_CONTRATADA_P" + i]);
                                workSheet.Cells[f, c].Style.Numberformat.Format = "#,##0";
                                c++;
                            }
                            else
                                c++;

                        }
                        else
                            c++;
                    }

                    if (r["TINDGCPY"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TINDGCPY"].ToString();
                    c++;

                    if (r["TTICONPS"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["TTICONPS"].ToString();
                    c++;

                    if (r["segmentoMercado"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["segmentoMercado"].ToString();
                    c++;

                    if (r["municipio"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["municipio"].ToString();
                    c++;

                    if (r["provincia"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["provincia"].ToString();
                    c++;

                    if (r["GESTION_PROPIA_ATR"] != System.DBNull.Value)
                    {
                        workSheet.Cells[f, c].Value = r["GESTION_PROPIA_ATR"].ToString();
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    c++;

                    if (r["TIPO_PUNTO_MEDIDA"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = Convert.ToInt32(r["TIPO_PUNTO_MEDIDA"]);
                    c++;

                    if (r["descripcion_autoconsumo"] != System.DBNull.Value)
                        workSheet.Cells[f, c].Value = r["descripcion_autoconsumo"].ToString();
                    c++;

                    if (r["F_ULTIMA_ACTUALIZACION"] != System.DBNull.Value)
                    {

                        workSheet.Cells[f, c].Value = Convert.ToDateTime(r["F_ULTIMA_ACTUALIZACION"]);
                        workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    c++;

                }
                db.CloseConnection();

                var allCells = workSheet.Cells[1, 1, f, c];
                headerCells = workSheet.Cells[1, 1, 1, c];
                headerFont = headerCells.Style.Font;
                headerFont.Bold = true;
                allCells = workSheet.Cells[1, 1, f, c];

                workSheet.View.FreezePanes(2, 1);
                workSheet.Cells["A1:AG1"].AutoFilter = true;
                allCells.AutoFitColumns();

                excelPackage.Save();
                
            }
            catch (Exception e)
            {
                
            }
        }

    }
}
