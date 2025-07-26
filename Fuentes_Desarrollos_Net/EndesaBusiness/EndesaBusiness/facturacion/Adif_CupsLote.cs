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
    public class Adif_CupsLote
    {
        public Int32 id_cups_lote { get; set; }
        public string cupsree { get; set; }
        public int lote { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }

        public void Save(Int32 id)
        {
            String strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "INSERT INTO adif_cups_lotes(id_cups_lote,cupsree,lote) values " +
                    "(" + id + ",'" + cupsree + "'," + lote + ");";
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
