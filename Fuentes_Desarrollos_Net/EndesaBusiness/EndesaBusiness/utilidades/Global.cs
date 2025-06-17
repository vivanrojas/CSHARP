using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class Global
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "Global");
        private Int32 applicationID;

        public Global()
        {
            applicationID = this.GetID();
        }


        public void CerrarVentana()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "update go_users set UltimoAcceso = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                    + " where UserCode = '" + System.Environment.UserName + "'";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("CerrarVentana: " + e.Message);
            }
        }

        private Int32 GetMaxID()
        {
            Int32 id = 0;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            string strSql;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                strSql = "Select max(ID) ID from global_applications;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    if (!reader.IsDBNull(0))
                    {
                        id = Convert.ToInt32(reader["ID"]);
                    }
                }


                db.CloseConnection();
                return id;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("GetMaxID " + e.Message);

            }
            return id;
        }

        public bool UsuarioValido()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            bool existe = false;
            try
            {
                
                strSql = "Select UserID from go_users where "
                + "UserCode =  '" + Environment.UserName + "';";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                existe = reader.Read();
                db.CloseConnection();
                return existe;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("UsuarioValido " + e.Message);
            }

            return false;
        }

        public string GetMailUser(string username)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string email = "";
            try
            {

                strSql = "Select UserEMail from go_users where "
                + "UserCode =  '" + username + "';";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["UserEMail"] != System.DBNull.Value)
                        email = r["UserEMail"].ToString();
                }
                db.CloseConnection();
                return email;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("GetMailUser " + e.Message);
                return null;
            }

            
        }

        private int GetID()
        {
            Int32 id = 0;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            string strSql;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                strSql = "Select ID from global_applications where "
                + "ApplicationName =  '" + System.AppDomain.CurrentDomain.FriendlyName + "';";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    id = Convert.ToInt32(reader["ID"]);
                }
                db.CloseConnection();
                if (id == 0)
                {
                    SaveID(GetMaxID() + 1);
                }
                else
                {
                    return id;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("GetID " + e.Message);
            }

            return 0;
        }

        private void SaveID(Int32 id)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                strSql = "REPLACE into global_applications SET " +
                    "ID = " + id + ", " +
                    "ApplicationName = '" + System.AppDomain.CurrentDomain.FriendlyName + "';";
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SaveID " + e.Message);
            }
        }


        public string GetParameter(string codigo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            string vcodigo;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                strSql = "Select Code from global_parameters where " +
                "ID = " + this.applicationID + " and " +
                "Codigo = '" + codigo + "';";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    vcodigo = reader["Code"].ToString();
                }
                else
                {
                    Console.WriteLine("El valor " + codigo + " no está parametrizado en global_parameters.");
                    vcodigo = "";
                }
                db.CloseConnection();
                return vcodigo;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("GetParameter " + e.Message);
            }
            return "";

        }


        public void SaveMD5(string description, DateTime lastLoad, DateTime startProcess,
            DateTime endProcess, string file, string md5, long fileSize)
        {

            MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                strSql = "REPLACE into global_md5 SET " +
                "ID = " + this.applicationID + ", " +
                "FileName = '" + file + "', " +
                "Description = '" + description + "', " +
                "LastLoad = '" + lastLoad.ToString("yyyy-MM-dd HH:mm:ss") + "', " +
                "StartProcess = '" + startProcess.ToString("yyyy-MM-dd HH:mm:ss") + "', " +
                "EndProcess = '" + endProcess.ToString("yyyy-MM-dd HH:mm:ss") + "', " +
                "MD5 = '" + md5 + "', " +
                "FileSize = " + fileSize + ";";
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SaveMD5 " + e.Message);
            }
        }

        public Boolean EqualMD5(string file, string md5)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            string vmd5 = "";
            try
            {
                Console.WriteLine("Buscando el MD5 para el archivo " + file);
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                strSql = "Select MD5 from global_md5 where " +
                "FileName = '" + file + "' order by LastLoad desc limit 1";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    vmd5 = reader.GetString(0);
                }
                db.CloseConnection();
                Console.WriteLine("Comparando " + md5 + " con " + vmd5);
                return vmd5.Equals(md5);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("EqualMD5 " + e.Message);
            }
            return false;
        }

        public void SaveQuery(String form, String query, DateTime begin, DateTime end)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {

                strSql = "replace into fo_process_performance set"
                    + " UserCode =  '" + Environment.UserName + "',"
                    + " Process = '" + form + "',"
                    + " Description = '" + query.Replace("'", "") + "',"
                    + " Begin = '" + begin.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + " End = '" + end.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + " TotalTime = " + (end - begin).TotalSeconds.ToString().Replace(",", ".") + ";";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SaveQuery " + e.Message);
            }
        }

        public void SaveProcess(string process, string description, DateTime processDate, DateTime begin, DateTime end)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;
            try
            {
                strSql = "REPLACE INTO global_process_performance SET"
                    + " ID = " + this.applicationID + ","
                    + " Process = '" + process + "',"
                    + " Description = '" + description + "',"
                    + " ProcessDate = '" + processDate.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + " Begin = '" + begin.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + " End = '" + end.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + " TotalTime = " + (end - begin).TotalSeconds.ToString().Replace(",", ".") + ";";

                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                ficheroLog.AddError("SaveProcess " + e.Message);
            }

        }

        public DateTime UltimaFechaCarga(string fichero)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            DateTime ultimaFecha = new DateTime();

            try
            {
                ultimaFecha = new DateTime(2010, 01, 01);

                strSql = "select LastLoad from global_md5_lastloads_v where"
                    + " FileName = '" + fichero + "' limit 1";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    ultimaFecha = Convert.ToDateTime(reader["LastLoad"]);
                }
                db.CloseConnection();

                return ultimaFecha;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("UltimaFechaCarga " + e.Message);
                return ultimaFecha;
            }
        }

        public List<String> GetSystemDriverList()
        {
            List<string> names = new List<string>();
            // get system dsn's
            Microsoft.Win32.RegistryKey reg = (Microsoft.Win32.Registry.LocalMachine).OpenSubKey("Software");
            if (reg != null)
            {
                reg = reg.OpenSubKey("ODBC");
                if (reg != null)
                {
                    reg = reg.OpenSubKey("ODBCINST.INI");
                    if (reg != null)
                    {

                        reg = reg.OpenSubKey("ODBC Drivers");
                        if (reg != null)
                        {
                            // Get all DSN entries defined in DSN_LOC_IN_REGISTRY.
                            foreach (string sName in reg.GetValueNames())
                            {
                                names.Add(sName);
                            }
                        }
                        try
                        {
                            reg.Close();
                        }
                        catch { /* ignore this exception if we couldn't close */ }
                    }
                }
            }

            return names;
        }
    }
}
