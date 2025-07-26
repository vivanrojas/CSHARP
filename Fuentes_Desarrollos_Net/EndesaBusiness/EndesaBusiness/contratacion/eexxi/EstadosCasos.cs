using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class EstadosCasos
    {
        public Dictionary<int, EndesaEntity.contratacion.xxi.Param_Estados_Casos> dic { get; set; }
        public EstadosCasos()
        {
            dic = new Dictionary<int, EndesaEntity.contratacion.xxi.Param_Estados_Casos>();
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
                strSql = "select estado_id, descripcion from cont.eexxi_param_estados_casos where"
                    + " (fromdate <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' and"
                    + " todate >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.xxi.Param_Estados_Casos c = new EndesaEntity.contratacion.xxi.Param_Estados_Casos();
                    if (r["estado_id"] != System.DBNull.Value)
                        c.estado_id = Convert.ToInt32(r["estado_id"]);
                    if (r["descripcion"] != System.DBNull.Value)
                        c.descripcion = r["descripcion"].ToString();

                    dic.Add(c.estado_id, c);

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "EstadosCasos - Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Warning);
            }
        }

    }
}
