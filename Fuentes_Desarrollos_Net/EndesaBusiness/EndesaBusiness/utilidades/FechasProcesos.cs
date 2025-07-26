using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.utilidades
{
    class FechasProcesos
    {
        Dictionary<string, EndesaEntity.global.Global_fechas_procesos> dic;

        public FechasProcesos()
        {
            dic = GetAll();
        }

        private Dictionary<string, EndesaEntity.global.Global_fechas_procesos> GetAll()
        {
            Dictionary<string, EndesaEntity.global.Global_fechas_procesos> d
                = new Dictionary<string, EndesaEntity.global.Global_fechas_procesos>();

            servidores.MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select area, proceso, descripcion, fecha, f_ult_mod"
                    + " from aux1.global_fechas_procesos";

                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.global.Global_fechas_procesos c = new EndesaEntity.global.Global_fechas_procesos();
                    c.area = r["area"].ToString();
                    c.proceso = r["proceso"].ToString();
                    c.descripcion = r["descripcion"].ToString();
                    c.fecha = Convert.ToDateTime(r["fecha"]);
                    c.f_ult_mod = Convert.ToDateTime(r["f_ult_mod"]);

                    EndesaEntity.global.Global_fechas_procesos o;
                    if (!d.TryGetValue(c.proceso, out o))                    
                        d.Add(c.proceso, c);
                    

                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public DateTime GetFechaProceso(string proceso)
        {
            EndesaEntity.global.Global_fechas_procesos o;
            if (dic.TryGetValue(proceso, out o))
                return o.fecha;
            else
                return DateTime.MinValue;
        }

        public void Update(string proceso, DateTime fecha)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;

            strSql = "update global_fechas_procesos set fecha = '" + fecha.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                + " where proceso = '" + proceso + "'";
            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

    }
}
