using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class GestionATRGas
    {
        utilidades.Param pp;
        logs.Log ficheroLog;
        public bool error_ftp;
        EndesaBusiness.utilidades.TelegramMensajes telegram;
        public GestionATRGas()
        {
            pp = new utilidades.Param("atrgas_param", servidores.MySQLDB.Esquemas.CON);
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_ATRGas");
            error_ftp = false;
            telegram = new EndesaBusiness.utilidades.TelegramMensajes(pp.GetValue("telegram_token_contratacion"), 
                pp.GetValue("telegram_channel_id"));
        }

        public void GuardaNumSecuencialTemporal(string secuencial_solicitud)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update atrgas_param set value = '" + secuencial_solicitud + "'"
                + " where code = 'secuencial_solicitud'";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            // Volvemos a cargar parametros para tener el ultimo valor en memoria
            pp = new utilidades.Param("atrgas_param", servidores.MySQLDB.Esquemas.CON);

        }

        public void Subir_a_FTP_Extremadura()
        {
            EndesaBusiness.utilidades.UltimateFTP ftp;
            string ruta_destino = "";
            string distribuidora = "";
            bool resultado_upload = false;

            try
            {
                string[] listaArchivos = Directory.GetFiles(pp.GetValue("inbox", DateTime.Now, DateTime.Now), "*.xml");
                if (listaArchivos.Length > 0)
                {

                    ficheroLog.Add("Conectando al FTP: " + pp.GetValue("ftp_extremadura_server", DateTime.Now, DateTime.Now));
                    ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        pp.GetValue("ftp_extremadura_server", DateTime.Now, DateTime.Now),
                        pp.GetValue("ftp_extremadura_user", DateTime.Now, DateTime.Now),
                        utilidades.FuncionesTexto.Decrypt(pp.GetValue("ftp_extremadura_pass", DateTime.Now, DateTime.Now), true),
                        pp.GetValue("ftp_extremadura_port", DateTime.Now, DateTime.Now));


                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        resultado_upload = false;
                        FileInfo fichero = new FileInfo(listaArchivos[i]);
                        ficheroLog.Add("Detectado: " + fichero.Name);
                        distribuidora = fichero.Name.Substring(16, 4);
                        ruta_destino = pp.GetValue("ftp_extremadura_ruta_destino", DateTime.Now, DateTime.Now);
                        //ftp.UploadInSecureFTP(pp.GetValue("ftp_extremadura_server", DateTime.Now, DateTime.Now) + fichero.Name, listaArchivos[i]);
                        resultado_upload = ftp.UploadFTPS(ruta_destino + fichero.Name, listaArchivos[i]);

                        if (resultado_upload)
                        {
                            if (pp.GetValue("utilizar_telegram") == "S")
                                telegram.SendMessage("Se publica en el FTP el archivo " + fichero.Name + " en la ruta " + ruta_destino);

                            ficheroLog.Add("Se publica en el FTP el archivo " + fichero.Name + " en la ruta " + ruta_destino);
                        }
                        else
                        {
                            if (pp.GetValue("utilizar_telegram") == "S")
                                telegram.SendMessage("CUIDADO ERROR!!!!"
                                   + System.Environment.NewLine
                                   + "================="
                                   + System.Environment.NewLine
                                   + "Necesario revisar proceso por ERROR en subida fichero FTPS Gas Extremadura");

                            ficheroLog.Add("ERROR en publicacion en el FTPS Gas Extremadura el archivo " + fichero.Name + " en la ruta " + ruta_destino);
                        }

                    }
                }
                error_ftp = false;
            }
            catch (Exception e)
            {
                error_ftp = true;
                ficheroLog.AddError("Subir_a_FTP_Extremadura: " + e.Message);

                if (pp.GetValue("utilizar_telegram") == "S")
                    telegram.SendMessage("CUIDADO ERROR!!!!"
                       + System.Environment.NewLine
                       + "================="
                       + System.Environment.NewLine
                       + "Necesario revisar proceso por ERROR: "
                       + System.Environment.NewLine
                       + e.Message);

                MessageBox.Show(e.Message,
                 "Error en conexión FTP",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);


            }
        }

        public void Subir_a_FTP()
        {
            EndesaBusiness.utilidades.UltimateFTP ftp;
            string ruta_destino = "";
            string distribuidora = "";

            try
            {
                string[] listaArchivos = Directory.GetFiles(pp.GetValue("inbox", DateTime.Now, DateTime.Now), "*.zip");
                if (listaArchivos.Length > 0)
                {

                    ficheroLog.Add("Conectando al FTP: " + pp.GetValue("ftp_sctd_server", DateTime.Now, DateTime.Now));
                    ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        pp.GetValue("ftp_sctd_server", DateTime.Now, DateTime.Now),
                        pp.GetValue("ftp_sctd_user", DateTime.Now, DateTime.Now),
                        utilidades.FuncionesTexto.Decrypt(pp.GetValue("ftp_sctd_pass", DateTime.Now, DateTime.Now), true),
                        pp.GetValue("ftp_sctd_port", DateTime.Now, DateTime.Now));


                    for (int i = 0; i < listaArchivos.Length; i++)
                    {
                        FileInfo fichero = new FileInfo(listaArchivos[i]);
                        ficheroLog.Add("Detectado: " + fichero.Name);
                        distribuidora = fichero.Name.Substring(16, 4);
                        ruta_destino = pp.GetValue("ftp_sctd_ruta_destino", DateTime.Now, DateTime.Now).Replace("distribuidora", distribuidora);
                        ftp.Upload(ruta_destino + fichero.Name, listaArchivos[i]);

                    }
                }
            }
            catch (Exception e)
            {
                ficheroLog.AddError("gestionATRGasFunciones.Subir_a_FTP: " + e.Message);
            }
        }

        public bool GuardaSolicitud(EndesaEntity.contratacion.gas.Solicitud s)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            int id = 0;
            bool firstOnly = true;

            try
            {

                id = NextIDRequest();

                sb.Append("replace into atrgas_solicitudes (id, mail_remitente, fecha_mail, asunto_mail, razon_social, nif,");
                sb.Append(" calle, numero, municipio, provincia, codigo_postal, cups, grupo_tarifario) values ");

                sb.Append("(").Append(id).Append(",");

                if (s.mail_remitente == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.mail_remitente).Append("',");

                if (s.fecha_mail == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.fecha_mail.ToString("yyyy-MM-dd HH:mm:ss")).Append("',");

                if (s.asunto_mail == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.asunto_mail).Append("',");

                if (s.razon_social == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.razon_social).Append("',");

                if (s.nif == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.nif).Append("',");

                if (s.calle == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.calle).Append("',");

                if (s.numero == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.numero).Append("',");

                if (s.municipio == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.municipio).Append("',");

                if (s.provincia == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.provincia).Append("',");

                if (s.codigo_postal == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.codigo_postal).Append("',");

                if (s.cups == null)
                    sb.Append("null,");
                else
                    sb.Append("'").Append(s.cups).Append("',");

                if (s.grupo_tarifario == null)
                    sb.Append("null)");
                else
                    sb.Append("'").Append(s.grupo_tarifario).Append("')");

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                sb.Clear();
                #region Guardamos los productos de la solicitud
                for (int i = 0; i < s.detalle.Count; i++)
                {
                    if (firstOnly)
                    {
                        sb.Append("replace into atrgas_solicitudes_detalle (id,linea,producto,qd,fecha_inicio,fecha_fin,qa, last_update_by) values ");
                        firstOnly = false;
                    }

                    sb.Append("(").Append(id).Append(",");
                    sb.Append(i + 1).Append(",");
                    sb.Append("'").Append(s.detalle[i].producto).Append("',");
                    sb.Append(s.detalle[i].qd.ToString().Replace(",", ".")).Append(",");
                    sb.Append("'").Append(s.detalle[i].fecha_inicio.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(s.detalle[i].fecha_fin.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append(s.detalle[i].qa.ToString().Replace(",", ".")).Append(",");
                    sb.Append("'").Append(System.Environment.UserName).Append("'),");

                }

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                #endregion

                sb = null;
                return false;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("GuardaSolicitud: " + e.Message);
                return true;
            }
        }

        private int NextIDRequest()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int maxID = 0;
            try
            {
                strSql = "select max(id) as id from atrgas_solicitudes;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["id"] != System.DBNull.Value)
                        maxID = Convert.ToInt32(r["id"]);
                }
                maxID++;
                db.CloseConnection();
                return maxID;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("UltimoIDSolicitud: " + e.Message);
                return 0;
            }
        }
    }
}
