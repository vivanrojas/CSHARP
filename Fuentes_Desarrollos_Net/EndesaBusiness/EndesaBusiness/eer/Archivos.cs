using EndesaBusiness.calendarios;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.eer
{
    public class Archivos : EndesaEntity.eer.Archivos
        
    {
        Dictionary<int, EndesaEntity.eer.Archivos> dic;
        public Archivos()
        {
            dic = Carga();
        }

        private Dictionary<int, EndesaEntity.eer.Archivos> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<int, EndesaEntity.eer.Archivos> d 
                = new Dictionary<int, EndesaEntity.eer.Archivos>();

            try
            {
                strSql = "select id_factura, ruta_factura, ruta_sstt,"
                    + " generado_pdf"                    
                    + " from eer_archivos order by id_factura";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.eer.Archivos c = new EndesaEntity.eer.Archivos();
                    if (r["id_factura"] != System.DBNull.Value)
                        c.id_factura = Convert.ToInt32(r["id_factura"]);
                    if (r["ruta_factura"] != System.DBNull.Value)
                        c.ruta_factura = r["ruta_factura"].ToString();
                    if (r["ruta_sstt"] != System.DBNull.Value)
                        c.ruta_sstt = r["ruta_sstt"].ToString();
                    if (r["generado_pdf"] != System.DBNull.Value)
                        c.generado_pdf = r["generado_pdf"].ToString() == "S";

                    EndesaEntity.eer.Archivos o;
                    if (!d.TryGetValue(c.id_factura, out o))
                        d.Add(c.id_factura, c);

                }
                db.CloseConnection();
                return d;
            }
            catch(Exception e)
            {
                return null;
            }

        }

        public void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update eer_archivos set "
                + " generado_pdf = 'S'"
                + " where id_factura = " + this.id_factura;
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }

        public void Insert()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "replace into eer_archivos (id_factura, ruta_factura, ruta_sstt) values "
                + "(" + this.id_factura + ","
                + "'" + this.ruta_factura + "',"
                + "'" + this.ruta_sstt + "')";
            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }
    }
}
