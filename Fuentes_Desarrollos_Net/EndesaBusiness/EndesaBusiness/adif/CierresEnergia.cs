using EndesaBusiness.calendarios;
using EndesaBusiness.servidores;
using EndesaEntity.contratacion.gas;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.adif
{
    public class CierresEnergia : EndesaEntity.medida.CierresEnergia
    {

        public Dictionary<string, List<EndesaEntity.medida.CierresEnergia>> dic;
        public CierresEnergia()
        {
            dic = new Dictionary<string, List<EndesaEntity.medida.CierresEnergia>>();
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
                strSql = "select cups20, fecha_desde, fecha_hasta"
                    + " from med.adif_cierres_energia";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.CierresEnergia c = new EndesaEntity.medida.CierresEnergia();

                    if (r["cups20"] != System.DBNull.Value)
                        c.cups20 = r["cups20"].ToString();

                    if (r["fecha_desde"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);

                    if (r["fecha_hasta"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);

                    List<EndesaEntity.medida.CierresEnergia> o;
                    if(!dic.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.medida.CierresEnergia>();
                        o.Add(c);
                        dic.Add(c.cups20, o);
                    }else
                        o.Add(c);   
                    

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "CierresEnergia - Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Warning);
            }
        }


        public bool ExisteCierre(string cups20, DateTime fd, DateTime fh)
        {
            bool existe = false;
            List<EndesaEntity.medida.CierresEnergia> o;
            if(dic.TryGetValue(cups20, out o))
            {
                List<EndesaEntity.medida.CierresEnergia> matches = o.Where(z => z.fecha_desde <= fd && z.fecha_hasta >= fh).ToList();
                existe = (matches.Count > 0);
            }
            return existe;

        }

        public void Save()
        {
            try
            {
                List<EndesaEntity.medida.CierresEnergia> o;
                if (!dic.TryGetValue(this.cups20, out o))
                    this.New();
                else if (dic.TryGetValue(this.cups20, out o))
                {
                    if (!o.Exists(z => z.cups20 == this.cups20))
                        this.New();
                    else
                        this.Update();
                }
                else
                    this.Update();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "CierresEnergia - Save",
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
                strSql = "insert into adif_cierres_energia (cups20, fecha_desde, fecha_hasta,"                    
                    + "  created_by, created_date) values ";

                if (this.cups20 != null)
                    strSql += "('" + this.cups20 + "'";
                else
                    strSql += "(null";

                if (this.fecha_desde != DateTime.MinValue)
                    strSql += ",'" + this.fecha_desde.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.fecha_hasta != DateTime.MinValue)
                    strSql += ",'" + this.fecha_hasta.ToString("yyyy-MM-dd") + "'";
                else
                    strSql += ",null";

                if (this.created_by != null)
                    strSql += ",'" + this.created_by + "'";
                else
                    strSql += ",'" + System.Environment.UserName + "'";

                strSql += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                MessageBox.Show("Se ha añadido correctamente el cierre para el CUPS: " + this.cups20,
                 "Contratos Gas",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information);

            }
            catch(Exception ex) 
            {
                MessageBox.Show(ex.Message,
             "CierresEnergia - New",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error);
            }
        }

        private void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            strSql = "update adif_cierres_energia set"
                + " last_update_by = '" + System.Environment.UserName.ToUpper() + "'";            

            if (this.fecha_desde != null)
                strSql += " ,fecha_desde = '" + this.fecha_desde.ToString("yyyy-MM-dd") + "'";

            if (this.fecha_hasta != null)
                strSql += " ,fecha_hasta = '" + this.fecha_hasta.ToString("yyyy-MM-dd") + "'";

            strSql += " where cups20 = '" + this.cups20 + "'";                

            db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

            MessageBox.Show("Se ha modifcado correctamente el cierre para el CUPS: " + this.cups20,
            "Cierres energía",
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
                    + "¿Desea borrar el cierre para el cups " + this.cups20 + "?"
                    , "Borrado del cierre " + this.cups20,
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {                   
                    strSql = "delete from adif_cierres_energia where cups20 = '" + this.cups20 + "'"
                        + " and fecha_desde = '" + this.fecha_desde.ToString("yyyy-MM-dd") + "'";

                    db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                    command = new MySqlCommand(strSql, db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();

                    MessageBox.Show("cierre borrado.",
                      "Cierres energía",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Information);

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "CierresEnergia - Del",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }
        }

    }
}
