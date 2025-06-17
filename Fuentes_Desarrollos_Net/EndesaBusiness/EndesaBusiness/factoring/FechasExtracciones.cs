using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.factoring
{
    public class FechasExtracciones
    {
        public FechasExtracciones() 
        {

        }

        public DateTime UltimaActualizacionObligCobradas()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            DateTime fecha = new DateTime();

            strSql = "SELECT MAX(f_ult_mod) max_fecha FROM fact.13_oblig_cob";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                fecha = Convert.ToDateTime(r["max_fecha"]);
            }
            db.CloseConnection();
            return fecha;

        }

        public DateTime UltimaActualizacionObligDeuda()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            DateTime fecha = new DateTime();

            strSql = "SELECT MAX(f_ult_mod) max_fecha FROM fact.deuda_obligaciones_original";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                fecha = Convert.ToDateTime(r["max_fecha"]);
            }
            db.CloseConnection();
            return fecha;

        }
    }
}
