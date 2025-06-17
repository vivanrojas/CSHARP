using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;

namespace EndesaBusiness.utilidades
{
    public class ParamUser: EndesaEntity.global.Parametro
    {
        private string vTabla;
        private MySQLDB.Esquemas vesquema;

        private List<EndesaEntity.global.Parametro> lista_parametros { get; set; }

        public ParamUser(string tabla, string userName, MySQLDB.Esquemas esquema)
        {
            lista_parametros = new List<EndesaEntity.global.Parametro>();
            vTabla = tabla;
            vesquema = esquema;
            this.user = userName.ToUpper();
            GetAll();
        }

        public void Save()
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();

            try
            {
                if (!Exist(code))
                {
                    sb.Append("insert into " + vTabla + " (code, user, from_date, to_date, value, description, created_by, created_date) values");
                    sb.Append(" ('").Append(code).Append("',");
                    sb.Append("'").Append(this.user).Append("',");
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
                    sb.Append(" set code = '" + code + "',");
                    sb.Append(" user = '" + this.user + "',");
                    sb.Append(from_date != DateTime.MinValue ? " ,from_date = '" + from_date.ToString("yyyy-MM-dd") + "'" : "");
                    sb.Append(to_date != DateTime.MinValue ? " ,to_date = '" + to_date.ToString("yyyy-MM-dd") + "'" : "");
                    sb.Append(value != null ? " ,value = '" + value + "'" : "");
                    sb.Append(description != null ? " ,description = '" + description + "'" : "");
                    sb.Append(" ,last_update_by = '" + System.Environment.UserName + "'");
                    sb.Append(" where code = '" + code + "' and");
                    sb.Append(" user = '" + this.user + "'");
                }



                db = new MySQLDB(vesquema);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
                GetAll();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public string GetValue(string code, DateTime fd, DateTime fh)
        {
            EndesaEntity.global.Parametro p = new EndesaEntity.global.Parametro();
            p = lista_parametros.Find(z => z.code == code && z.user.ToUpper() == this.user.ToUpper() && (z.from_date <= fd && z.to_date >= fh));

            if (p != null)
                return p.value;
            else
                return null;

        }


        private bool Exist(string code)
        {

            return lista_parametros.Exists(z => z.code == code && z.user == this.user);
        }

        public void Delete(string code, DateTime fd, DateTime fh)
        {
            string strSql = "";
            servidores.MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "delete from " + vTabla + " where"
                    + " code = '" + code + "' and"
                    + " user = '" + this.user + "' and"
                    + " from_date = '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " to_date = '" + fh.ToString("yyyy-MM-dd") + "';";
                db = new MySQLDB(vesquema);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();

                GetAll();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
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
                // lista_parametros = new List<Parametro>();
                strSql = "select * from " + vTabla
                    + " where user = '" + this.user + "'";
                db = new MySQLDB(vesquema);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.global.Parametro c = new EndesaEntity.global.Parametro();
                    c.code = reader["code"].ToString();
                    c.user = reader["user"].ToString().ToUpper();
                    c.from_date = Convert.ToDateTime(reader["from_date"]);
                    c.to_date = Convert.ToDateTime(reader["to_date"]);
                    c.value = reader["value"].ToString();
                    if (reader["description"] != System.DBNull.Value)
                        c.description = reader["description"].ToString();
                    if (reader["created_by"] != System.DBNull.Value)
                        c.created_by = reader["created_by"].ToString();
                    lista_parametros.Add(c);
                }
                reader.Close();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                Console.WriteLine("ParamUser.GetAll(): " + e.Message);
            }

        }

        private void GetAll(string user)
        {
            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            try
            {
                // lista_parametros = new List<Parametro>();
                strSql = "select * from " + vTabla
                    + " where user = '" + "ES02255021D" + "'";
                db = new MySQLDB(vesquema);
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    EndesaEntity.global.Parametro c = new EndesaEntity.global.Parametro();
                    c.code = reader["code"].ToString();
                    c.user = reader["user"].ToString().ToUpper();
                    c.from_date = Convert.ToDateTime(reader["from_date"]);
                    c.to_date = Convert.ToDateTime(reader["to_date"]);
                    c.value = reader["value"].ToString();
                    if (reader["description"] != System.DBNull.Value)
                        c.description = reader["description"].ToString();
                    if (reader["created_by"] != System.DBNull.Value)
                        c.created_by = reader["created_by"].ToString();
                    lista_parametros.Add(c);
                }
                reader.Close();
                
            }
            catch (Exception e)
            {
                Console.WriteLine("ParamUser.GetAll(): " + e.Message);
            }

            
        }

    }
}
