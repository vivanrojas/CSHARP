using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion
{
    public class TarifaDisuasoria
    {
        List<EndesaEntity.contratacion.Diasuario> lista_cups;
        public List<EndesaEntity.contratacion.Disuasorio_Resumen> lista_cups_resumen;
        public List<EndesaEntity.contratacion.Disuasoria_Informe> lista_Disusoria_informe;

        public TarifaDisuasoria()
        {
            lista_cups = new List<EndesaEntity.contratacion.Diasuario>();
            lista_cups_resumen = new List<EndesaEntity.contratacion.Disuasorio_Resumen>();
            lista_Disusoria_informe = new List<EndesaEntity.contratacion.Disuasoria_Informe>();
        }

        public string UltimaImportacion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql = "";
            
            string ult = null;

            try
            {
                strSql = "Select max(f_ult_mod) max from cont.c70ccnae";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    ult = reader["max"].ToString();
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return ult;
        }

        public string UltimaFechaPS_Historico()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;            
            string ult = null;

            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand("SELECT MAX(FECHA_ANEXION) max FROM cont.PS_AT_HIST", db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    ult = reader["max"].ToString();
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return ult;
        }

        //public DataTable CargadgvResumen(bool mayor50, DateTime fecha)
        //{

        //    MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader reader;
            

        //    try
        //    {

        //        db = new MySQLDB(MySQLDB.Esquemas.GBL);
        //        command = new MySqlCommand(C70CCNAE_CUPS_RESUMEN_HISTORICO(mayor50, fecha), db.con);

        //        MySqlDataAdapter da = new MySqlDataAdapter(command);
        //        DataTable dt = new DataTable();
        //        da.Fill(dt);
                

        //        reader = command.ExecuteReader();

        //        lista_cups_resumen = new List<EndesaEntity.contratacion.Disuasorio_Resumen>();
        //        while (reader.Read())
        //        {
        //            EndesaEntity.contratacion.Disuasorio_Resumen f = 
        //                new EndesaEntity.contratacion.Disuasorio_Resumen();
        //            f.num_cups = Convert.ToInt32(reader["NUM_CUPS"]);
        //            f.esencial = reader["esencial"].ToString();
        //            f.tipo_empresa_aapp = reader["TIPO_EMPRESA_AAPP"].ToString();
        //            lista_cups_resumen.Add(f);
        //        }

        //        if (dt.Rows.Count == 0)
        //        {
        //            MessageBox.Show("La consulta que ha realizado no devuelve ningún resultado.",
        //                "La consulta no devuelte datos.",
        //                MessageBoxButtons.OK,
        //                MessageBoxIcon.Information);
        //        }
        //        db.CloseConnection();
        //        return dt;

                
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message,
        //        "Error en la ejecución de la consulta C70CCNAE_CUPS_RESUMEN",
        //        MessageBoxButtons.OK,
        //        MessageBoxIcon.Error);

        //        return null;
        //    }

        //}


        public void CargadgvResumen()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            List<string> lista_nifs = new List<string>();
            
            lista_cups_resumen = new List<EndesaEntity.contratacion.Disuasorio_Resumen>();
            cartera.Cartera_SalesForce cartera;

            try
            {
                // Obtenemos los NIFs para consultar la cartera


                strSql = "select ps.NIF"
                   + " from cont.PS_AT ps"
                   + " left join cont.c70ccnae_cups_comp_exc des on des.cups13 = ps.IDU"
                   + " left join cont.c70ccnae cc on cc.cups13 = ps.IDU"
                   + " left join cont.cnae c on c.codigo = cc.cnae"
                   + " where des.cups13 is null and ps.EMPRESA = 'EEXXI'"
                   + " group by ps.NIF";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["NIF"] != System.DBNull.Value)
                        lista_nifs.Add(r["NIF"].ToString().Trim());
                }
                db.CloseConnection();

                cartera = new cartera.Cartera_SalesForce(lista_nifs);


                strSql = "select ps.fAltaCont AS FH_ALTA_CRTO,ps.IDU AS CD_CUPS,"
                    + " c.esencial AS esencial, ps.NIF,"
                    + " ps.VPOTCAL1 AS POTENCIA_MAX_CONTRATADA"
                    + " from cont.PS_AT ps"
                    + " left join cont.c70ccnae_cups_comp_exc des on des.cups13 = ps.IDU"
                    + " left join cont.c70ccnae cc on cc.cups13 = ps.IDU"
                    + " left join cont.cnae c on c.codigo = cc.cnae"
                    + " where des.cups13 is null and ps.EMPRESA = 'EEXXI'"
                    + " group by ps.NIF";                    
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
 
                    EndesaEntity.contratacion.Disuasorio_Resumen c = new EndesaEntity.contratacion.Disuasorio_Resumen();
                    

                    if (r["esencial"] != System.DBNull.Value)
                        c.esencial = r["esencial"].ToString();
                    else
                        c.esencial = "N";

                    switch (r["nif"].ToString().Substring(0, 1))
                    {
                        case "P":
                            c.tipo_empresa_aapp = "AAPP";
                            break;
                        case "Q":
                            c.tipo_empresa_aapp = "AAPP";
                            break;
                        case "S":
                            c.tipo_empresa_aapp = "AAPP";
                            break;
                        default:
                            cartera.GetCartera(r["nif"].ToString());                            
                            c.tipo_empresa_aapp = (cartera.segmento == "" || cartera.segmento == "NUCO") ? "Sin segmento" : cartera.segmento;
                            
                            break;
                    }

                    Agrupacion_Resumen(c.esencial, c.tipo_empresa_aapp);

                }
                db.CloseConnection();

            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la ejecución de la consulta C70CCNAE_CUPS_RESUMEN",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                
            }

        }


        private void Agrupacion_Resumen(string esencial, string tipo_empresa)
        {
            
            EndesaEntity.contratacion.Disuasoria_Informe c = new EndesaEntity.contratacion.Disuasoria_Informe();

            if (lista_Disusoria_informe.Count == 2)
            {
                
                for (int i = 0; i < lista_Disusoria_informe.Count; i++)
                {
                    if (lista_Disusoria_informe[i].etiqueta == "ESENCIAL" && esencial == "S")
                    {
                        switch (tipo_empresa)
                        {
                            case "KAM":
                                lista_Disusoria_informe[i].kam++;
                                break;
                            case "AAPP":
                                lista_Disusoria_informe[i].aapp++;
                                break;
                            case "Gestor":
                                lista_Disusoria_informe[i].gestor++;
                                break;
                            case "Sin segmento":
                                lista_Disusoria_informe[i].sin_segmento++;
                                break;
                        }



                        lista_Disusoria_informe[i].total_general++;
                    }

                    if (lista_Disusoria_informe[i].etiqueta != "ESENCIAL")
                    {
                        switch (tipo_empresa)
                        {
                            case "KAM":
                                lista_Disusoria_informe[i].kam++;
                                break;
                            case "AAPP":
                                lista_Disusoria_informe[i].aapp++;
                                break;
                            case "Gestor":
                                lista_Disusoria_informe[i].gestor++;
                                break;
                            default:
                                lista_Disusoria_informe[i].kam++;
                                break;
                        }                       

                        lista_Disusoria_informe[i].total_general++;
                    }


                }

            }
            else
            {
                #region Esencial
                c = new EndesaEntity.contratacion.Disuasoria_Informe();
                c.etiqueta = "ESENCIAL";                
                if(esencial == "S")
                {
                    switch (tipo_empresa)
                    {
                        case "KAM":
                            c.kam++;
                            break;
                        case "AAPP":
                            c.aapp++;
                            break;
                        case "Gestor":
                            c.gestor++;
                            break;
                    }
                    c.total_general++;
                }
                lista_Disusoria_informe.Add(c);

                #endregion
                #region Total_General
                c = new EndesaEntity.contratacion.Disuasoria_Informe();
                c.etiqueta = "Total general";
                if (esencial != "S")
                {
                    switch (tipo_empresa)
                    {
                        case "KAM":
                            c.kam++;
                            break;
                        case "AAPP":
                            c.aapp++;
                            break;
                        case "Gestor":
                            c.gestor++;
                            break;
                        default:
                            c.kam++;
                            break;

                    }
                    c.total_general++;
                }

                lista_Disusoria_informe.Add(c);
                #endregion


            }

           
        }

        public DataTable CargardgvCUPS(bool mayor50, DateTime fecha)
        {


            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;


            try
            {

                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(C70CCNAE_CUPS_HISTORICO(mayor50, fecha), db.con);

                MySqlDataAdapter da = new MySqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                reader = command.ExecuteReader();

                lista_cups = new List<EndesaEntity.contratacion.Diasuario>();

                while (reader.Read())
                {
                    EndesaEntity.contratacion.Diasuario f = new EndesaEntity.contratacion.Diasuario();
                    if (reader["FH_ALTA_CRTO"] != System.DBNull.Value)
                        f.fh_alta_contrato = Convert.ToDateTime(reader["FH_ALTA_CRTO"]);

                    f.cups13 = reader["CD_CUPS"].ToString();
                    f.esencial = reader["esencial"].ToString();
                    f.tipo_empresa = reader["TIPO_EMPRESA"].ToString();
                    f.aapp = reader["AAPP"].ToString();
                    f.tipo_empresa_aapp = reader["TIPO_EMPRESA_AAPP"].ToString();
                    f.potencia_max_contratada = Convert.ToDecimal(reader["POTENCIA_MAX_CONTRATADA"]);
                    lista_cups.Add(f);
                }
                db.CloseConnection();

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("La consulta que ha realizado no devuelve ningún resultado.",
                        "La consulta no devuelte datos.",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                return dt;
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message,
                "Error en la ejecución de la consulta C70CCNAE_CUPS",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;
            }

        }

        private string C70CCNAE_CUPS_HISTORICO(bool mayor50, DateTime f)
        {
            string strSql;

            strSql = "select if(ps.fAltaCont = 0, NULL, STR_TO_DATE(CONVERT(ps.fAltaCont,CHAR),'%Y%m%d')) AS FH_ALTA_CRTO, " +
                "ps.IDU AS CD_CUPS,if ((c.esencial = 'S'),'ESENCIAL','') AS esencial, " +
                "s.direccion AS TIPO_EMPRESA,if (((substr(ps.NIF, 1, 1) = 'P') or(substr(ps.NIF, 1, 1) = 'Q') or (substr(ps.NIF, 1, 1) = 'S')),'AAPP','GENERAL') AS AAPP," +
                "if ((if (((substr(ps.NIF, 1, 1) = 'P') or (substr(ps.NIF, 1, 1) = 'Q') or (substr(ps.NIF, 1, 1) = 'S')),'AAPP','GENERAL') = 'GENERAL'),s.direccion,'AAPP') AS TIPO_EMPRESA_AAPP," +
                "ps.VPOTCAL1 AS POTENCIA_MAX_CONTRATADA" +
                " from cont.PS_AT_HIST ps" +
                " left join cont.c70ccnae_cups_comp_exc des on des.cups13 = ps.IDU" +
                " left join cont.c70ccnae cc on cc.cups13 = ps.IDU" +
                " left join cont.cnae c on c.codigo = cc.cnae" +
                " left join cont.carteraSIOC s on s.cif = ps.NIF" +
                " where ps.Fecha_Anexion = '" + f.ToString("yyyy-MM-dd") + "' and" +
                " isnull(des.cups13) and ps.EMPRESA = 'EEXXI' group by ps.IDU order by ps.IDU;";

            return strSql;


        }




        private string C70CCNAE_CUPS_RESUMEN_HISTORICO(bool mayor50, DateTime f)
        {
            string strSql;

            strSql = "select count(*) NUM_CUPS, c.esencial, c.TIPO_EMPRESA_AAPP from";
            if (mayor50)
            {
                strSql += " cont.c70ccnae_cups_mayor50_hist as c";
                strSql += " where c.Fecha_Anexion = '" + f.ToString("yyyy-MM-dd") + "'";
            }
            else
            {
                strSql += " cont.c70ccnae_cups_menorigual50 as c";
            }

            strSql += " group by c.esencial, c.TIPO_EMPRESA_AAPP";

            return strSql;
        }

        public void ExportExcel(string fichero)
        {

            int f = 0;
            int c = 0;


            FileInfo fileInfo = new FileInfo(fichero);
            if (fileInfo.Exists)
                fileInfo.Delete();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("CASOS");

            var headerCells = workSheet.Cells[1, 1, 1, 6];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "Etiquetas de fila"; 
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "AAPP"; 
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "KAM"; 
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "GESTOR"; 
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "Sin segmento";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "Total General"; 
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            for (int i = 0; i < lista_Disusoria_informe.Count; i++)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = lista_Disusoria_informe[i].etiqueta;
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = lista_Disusoria_informe[i].aapp;
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = lista_Disusoria_informe[i].kam;
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = lista_Disusoria_informe[i].gestor;
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = lista_Disusoria_informe[i].sin_segmento;
                PintaRecuadro(excelPackage, f, c); c++;

                workSheet.Cells[f, c].Value = lista_Disusoria_informe[i].total_general;
                PintaRecuadro(excelPackage, f, c); c++;
            }

            var allCells = workSheet.Cells[1, 1, f, 5];
            workSheet.View.FreezePanes(2, 1);            
            allCells.AutoFitColumns();
            excelPackage.Save();

        }

        


        private void PintaRecuadro(ExcelPackage excelPackage, int f, int c)
        {
            var workSheet = excelPackage.Workbook.Worksheets.First();
            workSheet.Cells[f, c].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[f, c].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[f, c].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            workSheet.Cells[f, c].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        }
    }
}
