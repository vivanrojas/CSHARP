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
    public class TarifaATR
    {
        public Dictionary<string, EndesaEntity.Tarifa> dic { get; set; }
        public TarifaATR()
        {
            dic = new Dictionary<string, EndesaEntity.Tarifa>();
            CargaTarifas();
        }

        private void CargaTarifas()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select tarifa_atr, descripcion, eexxi from cont.eexxi_param_codigos_tarifas_atr;";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Tarifa c = new EndesaEntity.Tarifa();

                    if (r["tarifa_atr"] != System.DBNull.Value)
                        c.tarifa_ATR = r["tarifa_atr"].ToString(); ;

                    if (r["descripcion"] != System.DBNull.Value)
                        c.descripcion = r["descripcion"].ToString();

                    if (r["eexxi"] != System.DBNull.Value)
                        c.eexxi = r["eexxi"].ToString() == "S" ? true : false;

                    dic.Add(c.tarifa_ATR, c);

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "TarifaATR - CargaTarifas",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Warning);
            }
        }

        public string GetDescription(string tarifaATR)
        {
            EndesaEntity.Tarifa o;
            if (dic.TryGetValue(tarifaATR, out o))
                return o.descripcion;
            else
                return "";
        }

        public bool EsTarifaEEXXI(string tarifaATR)
        {
            EndesaEntity.Tarifa o;
            if (dic.TryGetValue(tarifaATR, out o))
                return o.eexxi;
            else
                return false;
        }
    }
}
