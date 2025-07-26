using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    public class UsageForm
    {

        public UsageForm()
        {

        }

        public void Start(string area, string form, string description)
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "replace into go_usage_form set"
                + " username = '" + System.Environment.UserName.ToUpper() + "',"
                + " area = '" + area + "',"
                + " form = '" + form + "',"
                + " description = '" + description + "',"
                + " start_date = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                + " end_date = null";
            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

        public void End(string area, string form, string description)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update go_usage_form set"
                + " end_date = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                + " where"
                + " username = '" + System.Environment.UserName.ToUpper() + "' AND"
                + " area = '" + area + "' AND"
                + " form = '" + form + "' AND"
                + " description = '" + description + "' AND"
                + " end_date is null";                
            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }

    }
}
