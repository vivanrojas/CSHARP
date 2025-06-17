using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class FechasProcesosPS
    {


        Dictionary<string, DateTime> dic;

        public FechasProcesosPS()
        {
            dic = Carga();
        }

        private Dictionary<string, DateTime> Carga()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, DateTime> d =
                new Dictionary<string, DateTime>();

            try
            {
                strSql = "SELECT proceso, f_ult_mod from ps_fechas_procesos";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    
                    if (r["proceso"] != System.DBNull.Value && r["f_ult_mod"] != System.DBNull.Value)
                    {
                        d.Add(r["proceso"].ToString(), Convert.ToDateTime(r["f_ult_mod"]));
                    }                       
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }


        }

        public DateTime FechaProceso(string proceso)
        {
            DateTime o;
            if (dic.TryGetValue(proceso, out o))
                return o;
            else
                return DateTime.MinValue;
        }

        public void ActualizaFechaProceso(string proceso)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            DateTime d = new DateTime();

            try
            {
                d = DateTime.Now;

                if (proceso == "PS" && ((int)d.DayOfWeek == 0 || (int)d.DayOfWeek == 6))
                {
                    if((int)d.DayOfWeek == 0)
                        d = d.AddDays(1);
                    else
                        d = d.AddDays(2);
                }
                    

                strSql = "update ps_fechas_procesos set"
                    + " fecha = '" + d.ToString("yyyy-MM-dd") + "'"
                    + " where proceso = '" + proceso + "'";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
