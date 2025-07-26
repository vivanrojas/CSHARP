using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.adif
{
    public class Provincias
    {
        public Dictionary<string, string> dic_provincias { get; set; }
        public Provincias() 
        {
            dic_provincias = Carga_Provincias();
        }

        private Dictionary<string, string> Carga_Provincias()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string provincia = "";
            string comunidad = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {
                strSql = "select comunidad, provincia from adif_provincias"
                    + " order by provincia";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    provincia = r["provincia"].ToString().ToUpper();
                    comunidad = r["comunidad"].ToString().ToUpper();
                    d.Add(provincia, comunidad);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public string GetComunidad(string provincia)
        {
            string o;
            if(dic_provincias.TryGetValue(provincia.ToUpper(), out o))
            { return o.ToUpper(); 
            }else
                return ""; 
            

        }

    }
}
