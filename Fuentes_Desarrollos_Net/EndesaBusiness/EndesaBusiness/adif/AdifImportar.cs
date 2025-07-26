using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.adif
{
    public class AdifImportar
    {
        public int TratarCarga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            MySQLDB db2;
            MySqlCommand command2;

            MySqlDataReader r;
            bool faltaInventario = false;
            int total_cups = 0;

            strSql = "SELECT t.cups20 from adif_medida_horaria_adif_temp t"
                + " LEFT OUTER JOIN adif_lotes l ON"
                + " l.CUPS20 = t.cups20"
                + " WHERE l.CUPS20 IS null";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                faltaInventario = true;
                // Movemos lo puntos que no pertenecen al inventario a la tabla adif_medida_horaria_adif_noinventario
                strSql = "replace into adif_medida_horaria_adif_noinventario "
                    + " select cups20 ,Fecha,TipoEnergia,Unidad,Total,Value1,Value2,Value3,"
                    + " Value4 , Value5, Value6, Value7, Value8, Value9, Value10, Value11,"
                    + " Value12, Value13, Value14, Value15, Value16, Value17, Value18,"
                    + " Value19, Value20, Value21, Value22, Value23, Value24, Value25,"
                    + " C1, C2, C3, C4, C5, C6, C7, C8, C9, C10, C11, C12, C13, C14, C15,"
                    + " C16, C17, C18, C19, C20, C21, C22, C23, C24, C25,"
                    + " FechaCarga, Fichero, EnConcentrador, FechaExportacion, User, F_ULT_MOD"
                    + " from adif_medida_horaria_adif_temp where cups20 = '" + r["cups20"].ToString() + "'";
                db2 = new MySQLDB(MySQLDB.Esquemas.MED);
                command2 = new MySqlCommand(strSql, db2.con);
                command2.ExecuteNonQuery();
                db2.CloseConnection();


                // Borramos los puntos de la tabla temp
                strSql = "delete from adif_medida_horaria_adif_temp where cups20 = '" + r["cups20"].ToString() + "'";
                db2 = new MySQLDB(MySQLDB.Esquemas.MED);
                command2 = new MySqlCommand(strSql, db2.con);
                command2.ExecuteNonQuery();
                db2.CloseConnection();
                             


                
            }
            db.CloseConnection();


            // Contamos todos los puntos tratados
            strSql = "select * from adif_medida_horaria_adif_temp a group by a.cups20;";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                total_cups++;
            }
            db.CloseConnection();


            // Movemos de temp a tabla definitiva
            strSql = "replace into adif_medida_horaria_adif select * from adif_medida_horaria_adif_temp;";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            // Movemos de temp a tabla definitiva
            strSql = "replace into adif_PO1011 select * from adif_PO1011_temp;";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "delete from adif_medida_horaria_adif_temp;";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            strSql = "delete from adif_PO1011_temp;";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            //strSql = "replace into adif_medida_horaria_adif_agrupada_vertical"
            //    + " select inv.LOTE, inv.TARIFA, med.FECHAHORA as FECHAHASTA,med.ESTACION,"
            //    + " sum(med.AE) A, sum(med.R1) R from adif_lotes inv inner join"
            //    + " (select FECHAHORA, ESTACION, ae, r1, CUPSREE, max(FechaCarga) as maximo"
            //    + " from adif_PO1011 med group by med.ESTACION, med.FECHAHORA, med.CUPSREE) med"
            //    + " on inv.CUPS20 = med.cupsree and"
            //    + " (inv.FECHA_DESDE <= med.FECHAHORA and inv.FECHA_HASTA >= med.FECHAHORA)"
            //    + " where med.maximo >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
            //    + " Group by"
            //    + " inv.LOTE, inv.Tarifa, med.ESTACION, med.FechaHora";
            //db = new MySQLDB(MySQLDB.Esquemas.MED);
            //command = new MySqlCommand(strSql, db.con);
            //command.ExecuteNonQuery();
            //db.CloseConnection();

            if (faltaInventario)
            {
                MessageBox.Show("Se ha encontrado puntos que no están en inventario."
                   + System.Environment.NewLine + System.Environment.NewLine
                   + "Por favor, revise la tabla adif_medida_horaria_adif_noinventario.",
                   "Puntos fuera de inventario",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
            }

            //ActualizaFechas();
            return total_cups;
        }
    }
}
