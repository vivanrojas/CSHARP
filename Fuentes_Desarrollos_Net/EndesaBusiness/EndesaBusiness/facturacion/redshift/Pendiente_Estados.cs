using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion.redshift
{
    public class Pendiente_Estados : EndesaEntity.medida.Pendiente
    {
        public Dictionary<string, EndesaEntity.medida.Pendiente> dic { get; set; }
        public Pendiente_Estados()
        {
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.medida.Pendiente> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string codigo = "";
            string descripcion = "";

            Dictionary<string, EndesaEntity.medida.Pendiente> d = new Dictionary<string, EndesaEntity.medida.Pendiente>();
            try
            {
                strSql = "SELECT cd_estado, de_estado FROM t_ed_p_estado_sap_pendiente_facturar";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Pendiente c = new EndesaEntity.medida.Pendiente();
                    c.cod_estado = r["cd_estado"].ToString();
                    c.descripcion_estado = r["de_estado"].ToString();
                    d.Add(c.cod_estado, c);
                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public void Save()
        {
            try
            {
                EndesaEntity.medida.Pendiente o;
                if (!dic.TryGetValue(this.cod_subestado, out o))
                    this.New();
                else
                    this.Update();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "Pendiente_Subestados - Save",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private void New()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                strSql = "insert into t_ed_p_estado_sap_pendiente_facturar (cd_estado, de_estado,"
                    + "  created_by, created_date) values ";

                if (this.cod_estado != null)
                    strSql += "('" + this.cod_estado + "'";
                else
                    strSql += "(null";

                if (this.descripcion_estado != null)
                    strSql += ",'" + this.descripcion_estado + "'";
                else
                    strSql += ",null";

                if (this.created_by != null)
                    strSql += ",'" + this.created_by + "'";
                else
                    strSql += ",'" + System.Environment.UserName + "'";

                strSql += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                db.MySQLTransaction();
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.MySQLCommit();
                db.CloseConnection();

                MessageBox.Show("Se ha añadido correctamente el registro con estado: " + this.cod_estado,
                 "Estados",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
             "Pendiente_Estados - New",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
            }
        }

        private void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update t_ed_p_estado_sap_pendiente_facturar set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

            if (this.cod_subestado != null)
                strSql += " ,cd_subestado = '" + this.cod_subestado + "'";

            if (this.descripcion_subestado != null)
                strSql += " ,de_subestado = '" + this.descripcion_subestado + "'";

            strSql += " where cd_subestado = '" + this.cod_subestado + "'";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            db.MySQLTransaction();
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.MySQLCommit();
            db.CloseConnection();

            MessageBox.Show("Se ha modifcado correctamente el subestado: " + this.cod_subestado,
            "Subestados",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        }

        public void Del()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            try
            {
                DialogResult result = MessageBox.Show("Warning:" + System.Environment.NewLine
                    + "¿Desea borrar el registro para el subestado " + this.cod_subestado + "?"
                    , "Borrado del registro con subestado " + this.cod_subestado,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    strSql = "delete from t_ed_p_estado_sap_pendiente_facturar where cd_subestado = '" + this.cod_subestado + "'";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("registro borrado.",
                      "Subestados",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "Pendiente_Subestados - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }
    }
}
