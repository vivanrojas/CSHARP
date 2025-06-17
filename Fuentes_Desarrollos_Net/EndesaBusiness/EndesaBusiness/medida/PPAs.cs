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

namespace EndesaBusiness.medida
{
    public class PPAs
    {
        utilidades.Param param;

        public int total_registros { get; set; }
        public int total_cups_meses { get; set; }

        List<string> lista_archivos { get; set; }


        public PPAs()
        {
            param = new utilidades.Param("ppas_param", MySQLDB.Esquemas.MED);
            

            if (!Directory.Exists(param.GetValue("ubicacion_salida_archivos", DateTime.Now, DateTime.Now)))
                Directory.CreateDirectory(param.GetValue("ubicacion_salida_archivos", DateTime.Now, DateTime.Now));

            if(param.GetValue("borrar_directorio", DateTime.Now, DateTime.Now) == "S")
            {
                string[] listaArchivos = Directory.GetFiles(param.GetValue("ubicacion_salida_archivos", DateTime.Now, DateTime.Now), "*.xlsx");

                if (listaArchivos.Length > 0)
                {
                    for (int j = 0; j < listaArchivos.Length; j++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[j]);
                        fichero.Delete();
                    }
                }
            }

            

        }


        public bool ExportarTabla_a_Excel()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, int> d = new Dictionary<string, int>();
            List<EndesaEntity.medida.PPAs_medida> l = new List<EndesaEntity.medida.PPAs_medida>();
            DateTime fd;
            DateTime fh;

            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            double percent = 0;
            int i = 0;
            string nombre_archivo = "";
            string[] cups20;

            try
            {

                lista_archivos = new List<string>();
                total_cups_meses = 0;

                strSql = "select count(*) total_registros"
                    + " from ppas_curvas_a_enviar pcae"
                    + " group by pcae.CUPS20, pcae.aaaammPdte";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                    total_cups_meses++;

                db.CloseConnection();


                pb.Text = "Generando archivos Excel";
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = total_cups_meses;


                strSql = "select pcae.CUPS20, pcae.aaaammPdte"
                    + " from ppas_curvas_a_enviar pcae"
                    + " group by pcae.CUPS20, pcae.aaaammPdte";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();

                while (r.Read())                
                    d.Add(r["CUPS20"].ToString() + "|" + Convert.ToInt32(r["aaaammPdte"]), Convert.ToInt32(r["aaaammPdte"]));

                db.CloseConnection();

                foreach(KeyValuePair<string, int> p in d)
                {
                    cups20 = p.Key.Split('|');
                    i++;
                    percent = (i / Convert.ToDouble(total_cups_meses)) * 100;
                    pb.progressBar.Increment(1);

                    fd = new DateTime();
                    fd = DateTime.MaxValue;
                    fh = new DateTime();
                    fh = DateTime.MinValue;


                    strSql = "SELECT NOMBRE, CUPS20, aaaammPdte, FECHA, HORA, AE, R1"
                        + " FROM med.ppas_curvas_a_enviar"
                        + " WHERE CUPS20 = '" + cups20[0] + "' and"
                        + " aaaammPdte = " + p.Value
                        + " order by CUPS20, FECHA, hora";
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();

                    while (r.Read())
                    {
                                            

                        EndesaEntity.medida.PPAs_medida c = new EndesaEntity.medida.PPAs_medida();
                        c.nombre = r["NOMBRE"].ToString();
                        c.cups20 = r["CUPS20"].ToString();
                        c.aaaammPdte = Convert.ToInt32(r["aaaammPdte"]);
                        c.fecha = Convert.ToDateTime(r["FECHA"]);
                        c.hora = Convert.ToInt32(r["HORA"]);
                        c.ae = Convert.ToInt32(r["AE"]);
                        c.r1 = Convert.ToInt32(r["R1"]);

                        fd = (fd > c.fecha ? c.fecha : fd);
                        fh = (fh < c.fecha ? c.fecha : fh);

                        l.Add(c);

                    }
                    db.CloseConnection();

                    nombre_archivo = param.GetValue("ubicacion_salida_archivos", DateTime.Now, DateTime.Now)
                        + cups20[0] + "_" + fd.ToString("yyyyMMdd") + "_" + fh.ToString("yyyyMMdd") + "_"
                        + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx";

                    lista_archivos.Add(nombre_archivo);

                    pb.txtDescripcion.Text = "Exportando " + nombre_archivo;
                    pb.Refresh();

                    Excel(nombre_archivo, cups20[0], fd, fh, l);             
                }
                pb.Close();

                return false;
            }
            catch(Exception e)
            {

                MessageBox.Show(e.Message,
                    "ExportarTabla_a_Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return true;
            }
        }

        public bool ExportarTabla_a_Excel_Agrupado()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<string, int> d = new Dictionary<string, int>();
            List<EndesaEntity.medida.PPAs_medida> l = new List<EndesaEntity.medida.PPAs_medida>();
            DateTime fd;
            DateTime fh;

            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            double percent = 0;
            int i = 0;
            string nombre_archivo = "";
            string[] cups20;

            try
            {

                fd = new DateTime();
                fd = DateTime.MaxValue;
                fh = new DateTime();
                fh = DateTime.MinValue;

                lista_archivos = new List<string>();
                total_cups_meses = 0;

                strSql = "select count(*) total_registros"
                    + " from ppas_curvas_a_enviar pcae"
                    + " group by pcae.CUPS20, pcae.aaaammPdte";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                    total_cups_meses++;

                db.CloseConnection();


                pb.Text = "Generando archivo Excel";
                pb.Show();
                pb.progressBar.Step = 1;
                pb.progressBar.Maximum = total_cups_meses;


                
                              
                
                strSql = "SELECT NOMBRE, CUPS20, aaaammPdte, FECHA, HORA, AE, R1"
                    + " FROM med.ppas_curvas_a_enviar"                    
                    + " order by CUPS20, FECHA, hora";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {


                    EndesaEntity.medida.PPAs_medida c = new EndesaEntity.medida.PPAs_medida();
                    c.nombre = r["NOMBRE"].ToString();
                    c.cups20 = r["CUPS20"].ToString();
                    c.aaaammPdte = Convert.ToInt32(r["aaaammPdte"]);
                    c.fecha = Convert.ToDateTime(r["FECHA"]);
                    c.hora = Convert.ToInt32(r["HORA"]);
                    c.ae = Convert.ToInt32(r["AE"]);
                    c.r1 = Convert.ToInt32(r["R1"]);

                    fd = (fd > c.fecha ? c.fecha : fd);
                    fh = (fh < c.fecha ? c.fecha : fh);

                    l.Add(c);

                }
                db.CloseConnection();

                nombre_archivo = param.GetValue("ubicacion_salida_archivos", DateTime.Now, DateTime.Now)
                    + "PPAs" + "_" + fd.ToString("yyyyMMdd") + "_" + fh.ToString("yyyyMMdd") + "_"
                    + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx";

                lista_archivos.Add(nombre_archivo);

                pb.txtDescripcion.Text = "Exportando " + nombre_archivo;
                pb.Refresh();

                Excel(nombre_archivo, null, fd, fh, l);
                
                pb.Close();

                return false;
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message,
                    "ExportarTabla_a_Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return true;
            }
        }

        private void Excel(string nombre_archivo, string cups20, DateTime fd, DateTime fh,
            List<EndesaEntity.medida.PPAs_medida> l)
        {

            int f = 0;
            int c = 0;

            

            FileInfo fileInfo = new FileInfo(nombre_archivo);



            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            var workSheet = excelPackage.Workbook.Worksheets.Add("PPAs");

            var headerCells = workSheet.Cells[1, 1, 1, 7];
            var headerFont = headerCells.Style.Font;
            f = 1;
            c = 1;
            headerFont.Bold = true;

            workSheet.Cells[f, c].Value = "NOMBRE"; c++;
            workSheet.Cells[f, c].Value = "CUPS20"; c++;
            workSheet.Cells[f, c].Value = "aaaammPdte"; c++;
            workSheet.Cells[f, c].Value = "FECHA"; c++;
            workSheet.Cells[f, c].Value = "HORA"; c++;
            workSheet.Cells[f, c].Value = "AE"; c++;
            workSheet.Cells[f, c].Value = "R1"; 

            for (int i = 1; i <= c; i++)
            {
                workSheet.Cells[f, i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells[f, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            foreach (EndesaEntity.medida.PPAs_medida p in l)
            {
                f++;
                c = 1;

                workSheet.Cells[f, c].Value = p.nombre; c++;
                workSheet.Cells[f, c].Value = p.cups20; c++;
                workSheet.Cells[f, c].Value = p.aaaammPdte; c++;

                if (p.fecha > DateTime.MinValue)
                {
                    workSheet.Cells[f, c].Value = p.fecha;
                    workSheet.Cells[f, c].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                c++;

                workSheet.Cells[f, c].Value = p.hora; c++;
                workSheet.Cells[f, c].Value = p.ae; c++;
                workSheet.Cells[f, c].Value = p.r1; c++;
               


            }

            var allCells = workSheet.Cells[1, 1, f, 7];
            workSheet.View.FreezePanes(2, 1);
            workSheet.Cells["A1:G1"].AutoFilter = true;
            allCells.AutoFitColumns();

            excelPackage.Save();

            
        }


        public bool SubidaFTP()
        {
            utilidades.UltimateFTP ftp;
            

            forms.FrmProgressBar pb = new forms.FrmProgressBar();
            double percent = 0;            

            try
            {                

                if (lista_archivos.Count > 0)                
                {

                    pb.Text = "Subiendo archivos a FTP " + param.GetValue("ftp_server", DateTime.Now, DateTime.Now);
                    pb.Show();
                    pb.progressBar.Step = 1;
                    pb.progressBar.Maximum = lista_archivos.Count;


                    ftp = new utilidades.UltimateFTP(
                    param.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                    param.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                    utilidades.FuncionesTexto.Decrypt(param.GetValue("ftp_pass", DateTime.Now, DateTime.Now), true),
                    param.GetValue("ftp_port", DateTime.Now, DateTime.Now));
            
                    for (int i = 0; i < lista_archivos.Count; i++)
                    {
                        FileInfo fichero = new FileInfo(lista_archivos[i]);

                        percent = (i+1 / Convert.ToDouble(lista_archivos.Count)) * 100;
                        pb.progressBar.Increment(1);
                        pb.txtDescripcion.Text = "Subiendo " + fichero.Name;
                        pb.Refresh();                        
                        ftp.Upload(param.GetValue("ruta_destino_FTP", DateTime.Now, DateTime.Now) + fichero.Name, lista_archivos[i]);
                        GuardaRegistro(fichero.Name.Substring(0, 20), fichero.Name);
                    }
                    pb.Close();                 

                    
                }
                return false;
            }
            catch(Exception e)
            {

                MessageBox.Show(e.Message,
                    "SubidaFTP",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                return true;

            }
        }

        private void GuardaRegistro(string cups, string nombre_archivo)
        {
            MySQLDB db;
            MySqlCommand command;            
            string strSql;

            strSql = "replace into ppas_control"
                   + " (cups20, archivo, fecha_exportacion, usuario) values"
                   + " ('" + cups + "',"
                   + " '" + nombre_archivo + "',"
                   + " '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                   + " '" + System.Environment.UserName + "')";


            db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
            
        }

        public DateTime UltimaExportacion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime d = new DateTime();


            strSql = "SELECT MAX(fecha_exportacion) as max_fecha FROM ppas_control";
            db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            if (r.Read())
                d = Convert.ToDateTime(r["max_fecha"]);
            db.CloseConnection();

            return d;

            db.CloseConnection();
        }

        public string Ayuda()
        {
            return param.GetValue("ayuda", DateTime.Now, DateTime.Now);
        }
    }
}
