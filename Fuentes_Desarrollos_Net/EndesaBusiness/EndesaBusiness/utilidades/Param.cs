using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace EndesaBusiness.utilidades
{
    public class Param: EndesaEntity.global.Parametro
    {
        private string vTabla;
        private MySQLDB.Esquemas vesquema;

        public List<EndesaEntity.global.Parametro> lista_parametros { get; set; }


        public Param(string tabla, MySQLDB.Esquemas esquema)
        {
            vTabla = tabla;
            vesquema = esquema;
            GetAll();
        }

        public void Save()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();

            try
            {
                if (!Exist(code, from_date.Date))
                {
                    sb.Append("insert into " + vTabla + " (code, from_date, to_date, value, description, created_by, created_date) values");
                    sb.Append(" ('").Append(code).Append("',");
                    sb.Append("'").Append(from_date.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(to_date.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(value).Append("',");
                    sb.Append("'").Append(description).Append("',");
                    sb.Append("'").Append(System.Environment.UserName).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("')");
                }
                else
                {
                    sb.Append("update ").Append(vTabla);
                    sb.Append(" set code = '" + code + "'");
                    sb.Append(from_date != DateTime.MinValue ? " ,from_date = '" + from_date.ToString("yyyy-MM-dd") + "'" : "");
                    sb.Append(to_date != DateTime.MinValue ? " ,to_date = '" + to_date.ToString("yyyy-MM-dd") + "'" : "");
                    sb.Append(value != null ? " ,value = '" + value + "'" : "");
                    sb.Append(description != null ? " ,description = '" + description + "'" : "");
                    sb.Append(" ,last_update_by = '" + System.Environment.UserName + "'");
                    sb.Append(" where code = '" + code + "'");
                }



                db = new MySQLDB(vesquema);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                GetAll();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "error en Param.Save",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public string GetValue(string code, DateTime fd, DateTime fh)
        {
            EndesaEntity.global.Parametro p = new EndesaEntity.global.Parametro();
            p = lista_parametros.Find(z => z.code == code && (z.from_date <= fd && z.to_date >= fh));

            if (p != null)
                return p.value;
            else
                return null;

        }

        public string GetValue(string code)
        {
            EndesaEntity.global.Parametro p = new EndesaEntity.global.Parametro();
            p = lista_parametros.Find(z => z.code == code && (z.from_date <= DateTime.Now.Date && z.to_date >= DateTime.Now.Date));

            if (p != null)
                return p.value;
            else
                return null;

        }

        public string GetDescription(string code)
        {
            EndesaEntity.global.Parametro p = new EndesaEntity.global.Parametro();
            p = lista_parametros.Find(z => z.code == code && (z.from_date <= DateTime.Now.Date && z.to_date >= DateTime.Now.Date));

            if (p != null)
                return p.description;
            else
                return null;

        }


        public bool ExisteParametroVigente(string code, DateTime fecha)
        {
            return lista_parametros.Exists(z => z.code == code && (z.from_date <= fecha.Date && z.to_date.Date >= fecha.Date));
        }

        private bool Exist(string code, DateTime fecha)
        {

            return lista_parametros.Exists(z => z.code == code && z.from_date == fecha);
        }

        public void SetCodeValueToDateTime(string code)
        {
            string strSql;
            servidores.MySQLDB db;
            MySqlCommand command;

            strSql = "update " + vTabla 
                + " SET VALUE = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'" 
                + " WHERE code = '"  + code + "'" 
                + " and from_date <= now() AND to_date >= now()";
            db = new MySQLDB(vesquema);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();

        }

        public void Delete(string code, DateTime fd, DateTime fh)
        {
            string strSql;
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                strSql = "delete from " + vTabla + " where"
                    + " code = '" + code + "' and"
                    + " from_date = '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " to_date = '" + fh.ToString("yyyy-MM-dd") + "';";
                db = new MySQLDB(vesquema);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();

                GetAll();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "No se ha podido borrar el registro:" + code,
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }


        private void GetAll()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;            
            try
            {



                lista_parametros = new List<EndesaEntity.global.Parametro>();


                strSql = "select * from " + vTabla;
                db = new MySQLDB(vesquema);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.global.Parametro c = new EndesaEntity.global.Parametro();
                    c.code = reader["code"].ToString();
                    c.from_date = Convert.ToDateTime(reader["from_date"]);
                    c.to_date = Convert.ToDateTime(reader["to_date"]);
                    c.value = reader["value"].ToString();

                    if(reader["description"] != System.DBNull.Value)
                        c.description = reader["description"].ToString();

                    if (reader["created_by"] != System.DBNull.Value)
                        c.created_by = reader["created_by"].ToString();

                    if (reader["last_update_date"] != System.DBNull.Value)
                        c.last_update_date = Convert.ToDateTime(reader["last_update_date"]);

                    if (reader["last_update_by"] != System.DBNull.Value)
                        c.last_update_by = reader["last_update_by"].ToString(); ;

                    lista_parametros.Add(c);
                }
                reader.Close();
                db.CloseConnection();
            

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "Error de parametrización.",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

        //public string GetParameter(servidores.MySQLDB.Esquemas esquema, string tabla, string codigo, DateTime fd, DateTime fh)
        //{
        //    servidores.MySQLDB db;
        //    MySqlCommand command;
        //    MySqlDataReader reader;
        //    string strSql;
        //    string vcodigo;
        //    try
        //    {
        //        db = new MySQLDB(esquema);
        //        strSql = "Select value from " + tabla + " where"
        //        + " code = '" + codigo + "' and"
        //        + " (from_date <= '" + fd.ToString("yyyy-MM-dd") + "' and"
        //        + " to_date >= '" + fh.ToString("yyyy-MM-dd") + "')";

        //        command = new MySqlCommand(strSql, db.con);
        //        reader = command.ExecuteReader();
        //        if (reader.Read())
        //        {
        //            vcodigo = reader["value"].ToString();
        //        }
        //        else
        //        {
        //            Console.WriteLine("El valor " + codigo + " no está parametrizado en " + tabla);
        //            vcodigo = "";
        //        }
        //        db.CloseConnection();
        //        return vcodigo;

        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message,
        //            "GetParameter " + esquema + " " + tabla + " " + codigo,
        //            MessageBoxButtons.OK,
        //            MessageBoxIcon.Error);
        //    }
        //    return "";

        //}

        public DateTime LastUpdateParameter(string parameter)
        {
            DateTime lastUpdate = new DateTime();
            lastUpdate = DateTime.MinValue;

            List<EndesaEntity.global.Parametro> lista = new List<EndesaEntity.global.Parametro>();
            lista = lista_parametros.FindAll(z => z.code == parameter);

            for (int i = 0; i < lista.Count; i++)
            {
                if (lista[i].last_update_date > lastUpdate)
                    lastUpdate = lista[i].last_update_date;
            }

            return lastUpdate;
        }

        public string GetLastUpdateBy(string parameter)
        {

            string update_by = "";

            List<EndesaEntity.global.Parametro> lista = new List<EndesaEntity.global.Parametro>();
            lista = lista_parametros.FindAll(z => z.code == parameter);

            for (int i = 0; i < lista.Count; i++)
            {
                update_by = lista[i].last_update_by;
            }

            return update_by;
        }

        public void UpdateParameter(string parameter, string value)
        {
                        
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update " + vTabla + " set value = '" + value + "',"
                + "last_update_by = '" + Environment.UserName.ToUpper() + "'"
                + " where code = '" + parameter + "'";
            db = new MySQLDB(vesquema);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();            
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
            var workSheet = excelPackage.Workbook.Worksheets.Add("Datos");

            var headerCells = workSheet.Cells[1, 1, 1, 4];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "Código";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "Desde fecha";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "Hasta fecha";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;

            workSheet.Cells[f, c].Value = "Valor";
            workSheet.Cells[f, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[f, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            workSheet.Cells[f, c].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            PintaRecuadro(excelPackage, f, c); c++;
                       

            foreach(EndesaEntity.global.Parametro p in lista_parametros)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.code; c++;
                workSheet.Cells[f, c].Value = p.from_date;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.to_date;
                workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern; c++;
                workSheet.Cells[f, c].Value = p.value; c++;



            }

            var allCells = workSheet.Cells[1, 1, f, 4];
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
