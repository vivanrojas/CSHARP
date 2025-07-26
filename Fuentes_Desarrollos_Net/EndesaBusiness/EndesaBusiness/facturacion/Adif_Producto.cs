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
    public class Adif_Producto
    {
        public Int32 id_producto { get; set; }
        public String producto { get; set; }

        public void Save(Int32 id)
        {
            String strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "INSERT INTO adif_productos (id_producto,producto) values " +
                    "(" + id + ",'" + producto + "');";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                 "Error a la hora de guardar producto.",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }
        }
    }
}
