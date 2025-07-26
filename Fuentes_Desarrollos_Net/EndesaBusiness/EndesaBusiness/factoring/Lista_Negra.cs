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
    public class Lista_Negra : EndesaEntity.factoring.ListaNegra
    {
        public Dictionary<string, EndesaEntity.factoring.ListaNegra> dic { get; set; }
        

        public Lista_Negra()
        {
            dic = new Dictionary<string, EndesaEntity.factoring.ListaNegra>();            
            Carga(GetQuery());
            
        }

        private string GetQuery()
        {
            string strSql = "";
            strSql = "select cnifdnic, dapersoc from ff_lista_negra";
            return strSql;
        }

        

        private void Carga(string query)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(query, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {

                if (r["cnifdnic"] != System.DBNull.Value)
                {
                    EndesaEntity.factoring.ListaNegra c = new EndesaEntity.factoring.ListaNegra();
                    c.nif = r["cnifdnic"].ToString();
                    c.cliente = r["dapersoc"].ToString();
                    EndesaEntity.factoring.ListaNegra o;
                    if (!dic.TryGetValue(c.nif, out o))
                        dic.Add(c.nif, c);
                }

            }
            db.CloseConnection();
        }

        

        public bool ExisteNIF(string nif)
        {

            EndesaEntity.factoring.ListaNegra o;
            return dic.TryGetValue(nif, out o);

        }

       

        public void Save()
        {
            try
            {
                EndesaEntity.factoring.ListaNegra o;
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
                strSql = "insert into ff_lista_negra (cnifdnic, dapersoc,"
                    + "  created_by, created_date) values ";

                if (this.nif != null)
                    strSql += "('" + this.nif + "'";
                else
                    strSql += "(null";

                if (this.cliente != null)
                    strSql += ",'" + this.cliente + "'";
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

                MessageBox.Show("Se ha añadido correctamente el registro con NIF: " + this.nif,
                 "Lista Negra",
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

            strSql = "update ff_lista_negra set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";

            if (this.nif != null)
                strSql += " ,cnifdnic = '" + this.nif + "'";

            if (this.cliente != null)
                strSql += " ,dapersoc = '" + this.cliente + "'";

            strSql += " where cnifdnic = '" + this.nif + "'";

            db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Se ha modifcado correctamente el nif: " + this.nif,
            "Lista Negra",
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
                    + "¿Desea borrar el registro para el nif " + this.nif + "?"
                    , "Borrado del registro con nif " + this.nif,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    strSql = "delete from ff_lista_negra where cnifdnic = '" + this.nif + "'";                        

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("registro borrado.",
                      "Lista Negra",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "Lista_NEgra - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }
    }
}
