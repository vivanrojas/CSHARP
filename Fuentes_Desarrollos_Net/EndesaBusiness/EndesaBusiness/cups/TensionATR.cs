using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.cups
{
    class TensionATR
    {
        Dictionary<int, EndesaEntity.contratacion.Tension> dic;
        public TensionATR()
        {
            dic = new Dictionary<int, EndesaEntity.contratacion.Tension>();
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
                strSql = "select codigo, descripcion from cont.eexxi_param_codigos_tensiones;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.Tension c = new EndesaEntity.contratacion.Tension();

                    if (r["codigo"] != System.DBNull.Value)
                        c.codigo = Convert.ToInt32(r["codigo"]);

                    if (r["descripcion"] != System.DBNull.Value)
                        c.descripcion = r["descripcion"].ToString();

                    dic.Add(c.codigo, c);

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "TensionATR - Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Warning);
            }
        }

        public string GetDescription(int codigo)
        {
            EndesaEntity.contratacion.Tension o;
            if (dic.TryGetValue(codigo, out o))
                return o.descripcion;
            else
                return "";
        }

    }
}
