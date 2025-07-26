using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    class SolicitudCurvasGestoresFuncionesDetalle : EndesaEntity.SolicitudDetalle
    {
        logs.Log ficheroLog;
        public SolicitudCurvasGestoresFuncionesDetalle()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CurvasDataMart_SolicitudDetalleFunciones");
        }

        public void Save()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";

            try
            {
                this.linea = GetNextLine(this.id);

                strSql = "replace into cc_sol_curvas_detalle (id, linea, cups20, fd, fh) values "
                    + "(" + this.id + ","
                    + this.linea + ","
                    + "'" + this.cups20 + "',"
                    + "'" + this.fd.ToString("yyyy-MM-dd") + "',"
                    + "'" + this.fh.ToString("yyyy-MM-dd") + "')";
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

        private int GetNextLine(int id)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            int linea = 1;

            try
            {
                strSql = "select max(linea) linea from cc_sol_curvas_detalle where"
                    + " id = " + id;
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["linea"] != System.DBNull.Value)
                    {
                        linea = Convert.ToInt32(r["linea"]) + 1;
                    }
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                ficheroLog.AddError(e.Message);
            }

            return linea;
        }
    }
}
