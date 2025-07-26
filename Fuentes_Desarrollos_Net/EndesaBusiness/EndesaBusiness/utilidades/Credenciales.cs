using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.utilidades
{
    public class Credenciales : EndesaEntity.global.Credencial
    {


        public Credenciales()
        {
            GetCredencial();
        }

        public Credenciales(string user)
        {
            GetCredencial(user);
        }

        private void GetCredencial()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            try
            {
                strSql = "select user_id, server_user, server_password " 
                    + " from mail_users_loging_info where "
                    + " user_id = '" + System.Environment.UserName + "' and"
                    + " (from_date <= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' and" 
                    + " to_date >= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    this.server_user = r["server_user"].ToString();
                    this.server_password = utilidades.FuncionesTexto.Decrypt(r["server_password"].ToString(),true);
                                    
                    
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "Error de parametrización en credenciales.",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }


        private void GetCredencial(string user)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            try
            {
                strSql = "select user_id, server_user, server_password "
                    + " from mail_users_loging_info where "
                    + " user_id = '" + user + "' and"
                    + " (from_date <= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' and"
                    + " to_date >= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    this.server_user = r["server_user"].ToString();
                    this.server_password = utilidades.FuncionesTexto.Decrypt(r["server_password"].ToString(), true);


                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "Error de parametrización en credenciales.",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }

        }

    }
}
