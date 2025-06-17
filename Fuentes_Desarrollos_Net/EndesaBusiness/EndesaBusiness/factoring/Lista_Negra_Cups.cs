using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.factoring
{
    public class Lista_Negra_Cups : EndesaEntity.factoring.ListaNegra_CUPS
    {
        public Dictionary<string, EndesaEntity.factoring.ListaNegra_CUPS> dic { get; set; }

        public Lista_Negra_Cups()
        {
            dic = new Dictionary<string, EndesaEntity.factoring.ListaNegra_CUPS>();
            Carga(GetQuery());
        }

        private string GetQuery()
        {
            string strSql = "";
            strSql = "select nif, cliente, cups20, fecha_inicio, fecha_fin, motivo from ff_lista_negra_cups";
            return strSql;
        }

        private void Carga(string query)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {



                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(query, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (r["cups20"] != System.DBNull.Value)
                    {
                        EndesaEntity.factoring.ListaNegra_CUPS c = new EndesaEntity.factoring.ListaNegra_CUPS();
                        if (r["nif"] != System.DBNull.Value)
                            c.nif = r["nif"].ToString();

                        if (r["cliente"] != System.DBNull.Value)
                            c.cliente = r["cliente"].ToString();

                        c.cups20 = r["cups20"].ToString();
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                        if (r["motivo"] != System.DBNull.Value)
                            c.motivo = r["motivo"].ToString();

                        EndesaEntity.factoring.ListaNegra_CUPS o;
                        if (!dic.TryGetValue(c.cups20, out o))
                            dic.Add(c.cups20, c);

                    }

                }
                db.CloseConnection();
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "Lista_Negra - Carga",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);
            }
        }


        public bool ExisteCUPS(string cups20)
        {

            EndesaEntity.factoring.ListaNegra_CUPS o;
            return dic.TryGetValue(cups20, out o);

        }

        public void Save()
        {
            try
            {
                EndesaEntity.factoring.ListaNegra_CUPS o;
                if (!dic.TryGetValue(this.nif, out o))
                    this.New();
                else
                    this.Update();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "Lista_Negra - Save",
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
                strSql = "insert into ff_lista_negra_cups (nif, cliente, cups20, fecha_inicio, fecha_fin,"
                    + " motivo, created_by, created_date) values ";

                if (this.nif != null)
                    strSql += "('" + this.nif + "'";
                else
                    strSql += "(null";

                if (this.cliente != null)
                    strSql += ",'" + this.cliente + "'";
                else
                    strSql += ",null";

                if (this.cups20 != null)
                    strSql += ",'" + this.cups20 + "'";
                else
                    strSql += ",null";

                if (this.fecha_inicio > DateTime.MinValue)
                    strSql += ",'" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.fecha_fin > DateTime.MinValue)
                    strSql += ",'" + this.fecha_fin.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.motivo != null)
                    strSql += ",'" + this.motivo + "'";
                else
                    strSql += ",null";

                if (this.created_by != null)
                    strSql += ",'" + this.created_by + "'";
                else
                    strSql += ",'" + System.Environment.UserName + "'";

                strSql += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                MessageBox.Show("Se ha añadido correctamente el registro con CUPS: " + this.cups20,
                 "Lista Negra CUPS",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
             "Lista Negra - New",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
            }
        }

        private void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update ff_lista_negra_cups set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

            if (this.nif != null)
                strSql += " ,nif = '" + this.nif + "'";

            if (this.cliente != null)
                strSql += " ,cliente = '" + this.cliente + "'";

            if (this.cups20 != null)
                strSql += " ,cups20 = '" + this.cups20 + "'";

            if (this.fecha_inicio > DateTime.MinValue)
                strSql += ",fecha_inicio = '" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'";            

            if (this.fecha_fin > DateTime.MinValue)
                strSql += ",fecha_fin = '" + this.fecha_fin.ToString("yyyy-MM-dd") + "'";            

            if (this.motivo != null)
                strSql += ",motivo = '" + this.motivo + "'";            

            strSql += " where cups20 = '" + this.cups20 + "' and"
                + " fecha_inicio = '" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Se ha añadido correctamente el registro con CUPS: " + this.cups20,
            "Lista Negra CUPS",
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
                    + "¿Desea borrar el registro para el cups " + this.cups20 + "?"
                    , "Borrado del registro con cups " + this.cups20,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    strSql = "delete from ff_lista_negra_cups where cups20 = '" + this.cups20 + "' and"
                        + " fecha_inicio = '" + this.fecha_inicio.ToString("yyyy-MM-dd") + "'";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("registro borrado.",
                      "Lista Negra CUPS",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "Lista_Negra - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }
    }
}
