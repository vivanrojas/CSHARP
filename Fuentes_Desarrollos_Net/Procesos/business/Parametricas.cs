using MySql.Data.MySqlClient;
using Procesos.servidores;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Procesos.business
{
    class Parametricas
    {
        public Dictionary<int, string> dic_procesos;
        public Dictionary<int, string> dic_periodicidad;
        public Dictionary<int, string> dic_tipoParametro;

        public Parametricas()
        {
            dic_procesos = CargaProcesos();
            dic_periodicidad = CargaPeriodicidad();
            dic_tipoParametro = CargaTipoParametro();

        }

        private Dictionary<int, string> CargaProcesos()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<int, string> d = new Dictionary<int, string>();
            int id = 0;
            string descripcion = "";


            try
            {
                strSql = "select id_p_proceso, descripcion from aux1.q_p_procesos;";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["id_p_proceso"] != System.DBNull.Value)
                        id = Convert.ToInt32(r["id_p_proceso"]);

                    if (r["descripcion"] != System.DBNull.Value)
                        descripcion = r["descripcion"].ToString();

                    string o;
                    if (!d.TryGetValue(id, out o))
                        d.Add(id, descripcion);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private Dictionary<int, string> CargaPeriodicidad()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<int, string> d = new Dictionary<int, string>();
            int id = 0;
            string descripcion = "";


            try
            {
                strSql = "select id_p_periodicidad, descripcion from aux1.q_p_periodicidad;";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["id_p_periodicidad"] != System.DBNull.Value)
                        id = Convert.ToInt32(r["id_p_periodicidad"]);

                    if (r["descripcion"] != System.DBNull.Value)
                        descripcion = r["descripcion"].ToString();

                    string o;
                    if (!d.TryGetValue(id, out o))
                        d.Add(id, descripcion);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private Dictionary<int, string> CargaTipoParametro()
        {

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            Dictionary<int, string> d = new Dictionary<int, string>();
            int id = 0;
            string descripcion = "";


            try
            {
                strSql = "select id_p_parametro, descripcion from aux1.q_p_parametros;";
                db = new MySQLDB(MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["id_p_parametro"] != System.DBNull.Value)
                        id = Convert.ToInt32(r["id_p_parametro"]);

                    if (r["descripcion"] != System.DBNull.Value)
                        descripcion = r["descripcion"].ToString();

                    string o;
                    if (!d.TryGetValue(id, out o))
                        d.Add(id, descripcion);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public string GetProceso(int id)
        {
            string o;
            if (dic_procesos.TryGetValue(id, out o))
                return o;
            else
                return "";
        }

        public string GetPeriodicidad(int id)
        {
            string o;
            if (dic_periodicidad.TryGetValue(id, out o))
                return o;
            else
                return "";
        }

        public string GetParametro(int id)
        {
            string o;
            if (dic_tipoParametro.TryGetValue(id, out o))
                return o;
            else
                return "";
        }

        public int GetIDProceso(string proceso)
        {
            foreach (KeyValuePair<int, string> p in dic_procesos)
                if (p.Value == proceso)
                    return p.Key;

            return 0;
        }

        public int GetIDPeriodicidad(string periodicidad)
        {
            foreach (KeyValuePair<int, string> p in dic_periodicidad)
                if (p.Value == periodicidad)
                    return p.Key;

            return 0;
        }

        public int GetIDParametro(string parametro)
        {
            foreach (KeyValuePair<int, string> p in dic_tipoParametro)
                if (p.Value == parametro)
                    return p.Key;

            return 0;
        }

    }
}
