using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.adif
{
    public class Param
    {
        public Param()
        {

        }

        public string GetParam(string codigo)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            strSql = "SELECT valor FROM adif_parametros WHERE codigo = '" + codigo + "'";
            db = new MySQLDB(MySQLDB.Esquemas.MED);

            command = new MySqlCommand(strSql, db.con);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                strSql = reader["valor"].ToString();
            }
            db.CloseConnection();
            return strSql;
        }
    }
}
