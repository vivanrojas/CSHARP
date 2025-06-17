using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.adif
{
    public class AdifLotes
    {

        public List<string> lista_lotes { get; set; }
        public List<string> lista_CUPS13 { get; set; }
        public AdifLotes(DateTime fd, DateTime fh)
        {
            lista_lotes = CargaLotes(fd, fh);
            lista_CUPS13 = CargaCUPS13(fd, fh);
        }

        private List<string> CargaLotes(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            List<string> l = new List<string>();

            try
            {
                strSql = "select l.LOTE from med.adif_lotes l"
                    + " where (l.FECHA_DESDE <= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " and l.FECHA_HASTA >= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " group by l.LOTE;";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    l.Add(r["LOTE"].ToString());
                }
                db.CloseConnection();
                return l;
            }
            catch(Exception e)
            {
                return null;
            }
        }

        private List<string> CargaCUPS13(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            List<string> l = new List<string>();

            try
            {
                strSql = "select CUPS13 from adif_lotes_v v where"
                + " v.FECHA_DESDE <= '" + fd.ToString("yyyy-MM-dd") + "' and "
                + " v.FECHA_HASTA >= '" + fh.ToString("yyyy-MM-dd") + "';";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    l.Add(r["CUPS13"].ToString());
                }
                db.CloseConnection();
                return l;
            }
            catch (Exception e)
            {
                return null;
            }
        }


    }
}
