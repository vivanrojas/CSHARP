using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida.curvasDataMart
{
    public class SolicitudFunciones : EndesaEntity.Solicitud
    {
        logs.Log ficheroLog;
        public SolicitudFunciones()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CurvasDataMart_SolicitudFunciones");
        }

        public void Save()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                this.id = GetNextID();
                strSql = "replace into cc_sol_curvas (id, mail, fechahora_mail, desc_error) values "
                    + "(" + this.id + ","
                    + "'" + this.mail + "',"
                    + "'" + this.fechahora_mail.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + "'" + this.desc_error + "')";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }
        }

        private int GetNextID()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int id = 1;

            try
            {
                strSql = "select max(id) id from cc_sol_curvas";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["id"] != System.DBNull.Value)
                    {
                        id = Convert.ToInt32(r["id"]) + 1;
                    }
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }

            return id;
        }

        private int GetLastID()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int id = 1;

            try
            {
                strSql = "select max(id) id from cc_sol_curvas";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["id"] != System.DBNull.Value)
                    {
                        id = Convert.ToInt32(r["id"]);
                    }
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }

            return id;
        }
    }
}
