using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.adif
{
    public class ProcesosFunciones
    {
        public ProcesosFunciones()
        {

        }

        public void SaveProcess(string process, string description, DateTime begin, DateTime end)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "REPLACE into adif_process_performance SET "
                + "Process = '" + process + "', "
                + "description = '" + description + "', "
                + "Begin = '" + begin.ToString("yyyy-MM-dd") + "', "
                + "End = '" + end.ToString("yyyy-MM-dd") + "', "
                + "TotalTime = " + (end - begin).TotalSeconds.ToString().Replace(",", ".") + ", "
                + "user = '" + System.Environment.UserName + "';";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public DateTime GetLastProcess(string process)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime ultimo = new DateTime();

            ultimo = new DateTime(2001, 01, 01);
            strSql = "select max(begin) begin from adif_process_performance where"
                + " process = '" + process + "';";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                ultimo = Convert.ToDateTime(r["begin"]);
            }

            return ultimo;
        }
    }
}
