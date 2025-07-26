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
    public class FechasCargas
    {
        public static DateTime UltimaActualizacionPS(string nombreProceso)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();

            try
            {
                strSql = "select f_ult_mod from ps_fechas_procesos where"
                    + " proceso = '" + nombreProceso + "'";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["f_ult_mod"] != System.DBNull.Value)
                        fecha = Convert.ToDateTime(r["f_ult_mod"]);
                }

                db.CloseConnection();
                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "UltimaActualizacionCuadroMando: " + nombreProceso,
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1999, 01, 01);

            }
        }

        public static DateTime UltimaSolicitud_EEXXI_XML()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;
            DateTime fecha = new DateTime();
            bool encontrado = false;

            try
            {

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                strSql = "select max(t.FechaSolicitud) as max_fecha from eexxi_solicitudes_tmp t;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (reader["max_fecha"] != System.DBNull.Value)
                    {
                        fecha = Convert.ToDateTime(reader["max_fecha"]);
                        encontrado = true;
                    }

                }
                db.CloseConnection();

                if (!encontrado)
                {
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    strSql = "select max(t.FechaSolicitud) as max_fecha from eexxi_solicitudes t;";
                    command = new MySqlCommand(strSql, db.con);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        if (reader["max_fecha"] != System.DBNull.Value)
                            fecha = Convert.ToDateTime(reader["max_fecha"]);
                    }

                    db.CloseConnection();
                }


                return fecha;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
             "FechasCargas -- >UltimaSolicitud_EEXXI_XML: ",
             MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return new DateTime(1999, 01, 01);

            }
        }
    }
}
