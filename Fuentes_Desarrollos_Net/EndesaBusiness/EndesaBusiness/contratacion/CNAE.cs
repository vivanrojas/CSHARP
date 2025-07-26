using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion
{
    class CNAE
    {
        public string codigo { get; set; }
        public string codintegr { get; set; }
        public string descripcion { get; set; }
        public bool esencial { get; set; }

        Dictionary<string, EndesaEntity.contratacion.CNAE_Tabla> dic;

        public CNAE()
        {
            dic = new Dictionary<string, EndesaEntity.contratacion.CNAE_Tabla>();
            Carga();
        }


        private void Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select codigo, codintegr, descripcion, esencial from cnae";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.CNAE_Tabla c = new EndesaEntity.contratacion.CNAE_Tabla();

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "Error en la carga de facturas",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }

        }

        public void Add()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            utilidades.Global g = new utilidades.Global();

            try
            {
                strSql = "replace into cnae (codigo, codintegr, descripcion, esencial, user) values"
                    + " ('" + codigo + "','" + codintegr + "','" + descripcion + "','"
                    + (esencial ? "S" : "N") + "','" + Environment.UserName + "');";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormCNAE", strSql, DateTime.Now, DateTime.Now);

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
                strSql = "delete from cnae where codigo = '" + codigo + "';";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormCNAE", strSql, DateTime.Now, DateTime.Now);

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
                strSql = "update cnae set esencial = '" + (esencial ? "S" : "N") + "'";

                if (codintegr != null)
                {
                    strSql = strSql + " ,codintegr = '" + codintegr + "'";
                }
                if (descripcion != null)
                {
                    strSql = strSql + " ,descripcion = '" + descripcion + "'";
                }

                strSql = strSql + " ,user = '" + Environment.UserName + "'";

                strSql = strSql + " where codigo = '" + codigo + "'";


                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormCNAE", strSql, DateTime.Now, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error in data deletion",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }
    }
}
