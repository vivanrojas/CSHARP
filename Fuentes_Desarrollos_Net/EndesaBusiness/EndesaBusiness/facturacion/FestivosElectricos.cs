using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class FestivosElectricos
    {
        public string tabla { get; set; }
        public DateTime fechaFestivo { get; set; }
        public string descripcion { get; set; }

        public void Add()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            utilidades.Global g = new utilidades.Global();

            try
            {
                strSql = "replace into " + tabla + " (FechaFestivo, Descripcion, USER) values"
                    + " ('" + fechaFestivo.ToString("yyyy-MM-dd") + "',"
                    + (descripcion != null ? " '" + descripcion + "'" : " null") + ",'" + Environment.UserName + "');";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormFacFestivosElectricosParam", strSql, DateTime.Now, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el guardado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public void Del()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            utilidades.Global g = new utilidades.Global();


            try
            {
                strSql = "delete from " + tabla + " where FechaFestivo = '" + fechaFestivo.ToString("yyyy-MM-dd") + "';";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormFacFestivosElectricosParam", strSql, DateTime.Now, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el borrado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            utilidades.Global g = new utilidades.Global();

            try
            {
                strSql = "update " + tabla + " set Descripcion = '" + descripcion + "'";
                strSql = strSql + " ,user = '" + Environment.UserName + "'";

                strSql = strSql + " where FechaFestivo = '" + fechaFestivo.ToString("yyyy-MM-dd") + "'";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormFacFestivosElectricosParam", strSql, DateTime.Now, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el guardado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }
    }
}
