using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.punto_suministro
{
    public class Territorios
    {
        public Dictionary<int, string> dic { get; set; }
        Dictionary<string, List<string>> dic_territorios { get; set; }

        public Territorios()
        {
            dic = Carga();
            dic_territorios = CargaTerritorios();
        }

        private Dictionary<int, string> Carga()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<int, string> d = new Dictionary<int, string>();

            try
            {
                strSql = "SELECT Territorio, Provincia, CodigoPostal"
                + " FROM fact.fact_territorios";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    d.Add(Convert.ToInt32(r["CodigoPostal"]), r["Territorio"].ToString());
                }
                db.CloseConnection();

                return d;

            }catch(Exception e)
            {
                MessageBox.Show(e.Message, "Territorios.Carga", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
        }

        private Dictionary<string, List<string>> CargaTerritorios()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, List<string>> d
                = new Dictionary<string, List<string>>();

            try
            {
                strSql = "SELECT Tarifa, Territorio from fact_calendarios"
                    + " group by Tarifa, Territorio";                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    List<string> o;
                    if (!d.TryGetValue(r["Tarifa"].ToString(), out o))
                    {
                        o = new List<string>();
                        o.Add(r["Territorio"].ToString()); 
                        d.Add(r["Tarifa"].ToString(), o);
                    }
                    else
                        o.Add(r["Territorio"].ToString());
                }
                db.CloseConnection();

                return d;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Territorios.CargaTerritorios", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
        }

        public string GetTerritorio(string codigoPostal)
        {
            string o;
            int codigo = Convert.ToInt32(codigoPostal.Substring(0, 2));

            if (dic.TryGetValue(codigo, out o))
                return o;
            else
            {
                MessageBox.Show("No se ha podido encontrar el territorio para el "
                    + "codigo postal " + codigoPostal,
                    "Territorios.GetTerritorio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
                
        }

        public List<string> GetTerritorios(string tarifa)
        {
            List<string> o;
            string _tarifa = "";

            if(tarifa.Contains("TD"))
                switch (Convert.ToInt32(tarifa.Substring(0,1)))
                {
                    case 2:
                        _tarifa = "2.0TD";
                        break;
                    case 3:
                        _tarifa = "3.0TD";
                        break;
                    case 6:
                        _tarifa = tarifa;
                        break;
                }
            else if (tarifa.Contains("VE"))
                switch (Convert.ToInt32(tarifa.Substring(0, 1)))
                {
                    case 3:
                        _tarifa = "3.0TDVE";
                        break;
                    case 6:
                        _tarifa = "6.1TDVE";
                        break;
                }
            else
                switch (Convert.ToInt32(tarifa.Substring(0, 1)))
                {
                    case 2:
                        _tarifa = tarifa;
                        break;
                    case 3:
                        _tarifa = tarifa;
                        break;
                    case 6:
                        _tarifa = "6.X";
                        break;
                }


            if (dic_territorios.TryGetValue(_tarifa, out o))
                return o;
            else
                return null;
        }

    }
}
