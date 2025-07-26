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
    public class Adif_Empresa
    {
        public Int32 id_empresa { get; set; }
        public string empresa { get; set; }

        public void Save(Int32 id)
        {
            String strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "INSERT INTO adif_empresas(id_empresa,empresa) values " +
                    "(" + id + ",'" + empresa + "');";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                 "Error a la hora de guardar cups-lote.",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }
        }
    }
}
