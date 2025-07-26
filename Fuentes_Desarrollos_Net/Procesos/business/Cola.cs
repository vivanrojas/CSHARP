using MySql.Data.MySqlClient;
using Procesos.servidores;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Procesos.business
{
    class Cola : EndesaEntity.cola.Cola
    {
        public Dictionary<string, EndesaEntity.cola.Cola> dic { get; set; }
        public Cola()
        {
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.cola.Cola> Carga()
        {
            Dictionary<string, EndesaEntity.cola.Cola> d =
                new Dictionary<string, EndesaEntity.cola.Cola>();

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select cola, descripcion, mail_aviso from q_colas";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.AUX);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cola.Cola c = new EndesaEntity.cola.Cola();
                    c.cola = r["cola"].ToString();
                    c.descripcion = r["descripcion"].ToString();
                    c.mail_aviso = r["mail_aviso"].ToString();

                    d.Add(c.cola, c);

                }
                db.CloseConnection();

                return d;
            }
            catch(Exception e)
            {
                return null;
            }

        }

        public void Save()
        {
            New();
        }

        private void New()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                if (ExisteCola())
                {
                    MessageBox.Show("Ya existe una cola de procesos con este nombre."
                        + System.Environment.NewLine
                        + "Por favor, seleccione otro nombre.",
                    "Nueva Cola de procesos",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Error);
                }
                else
                {
                    strSql = "insert into q_colas (cola, descripcion, mail_aviso, usuario, f_ult_mod) values ";
                    strSql += "('" + this.cola + "'";
                    strSql += ",'" + this.descripcion + "'";
                    strSql += ",'" + this.mail_aviso + "'";
                    strSql += ",'" + System.Environment.UserName + "'";
                    strSql += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.AUX);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();

                    MessageBox.Show("La cola se ha creado correctamente.",
                    "Nueva Cola de procesos",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);
                }
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "Cola - New",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Error);
            }
        }

        public bool ExisteCola()
        {
            EndesaEntity.cola.Cola o;
            return (dic.TryGetValue(this.cola, out o));
        }

        public string UltimaColaCreada()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            strSql = "select cola, max(f_ult_mod) f_ult_mod from q_colas"
                + " where usuario = '" + System.Environment.UserName + "'";
            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
                return r["cola"].ToString();

            return "";
        }

        public bool HayNuevaCola()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            DateTime fecha = new DateTime();

            strSql = "select cola, max(f_ult_mod) f_ult_mod from q_colas"
                + " where usuario = '" + System.Environment.UserName + "'";
            db = new MySQLDB(MySQLDB.Esquemas.AUX);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (r["f_ult_mod"] != System.DBNull.Value)
                    fecha = Convert.ToDateTime(r["f_ult_mod"]);
                else
                    fecha = DateTime.MinValue;
            }
                

            return (fecha.AddMinutes(2) > DateTime.Now);
                
        }

    }
}
