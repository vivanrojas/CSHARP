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
    public class CupsFunciones : EndesaEntity.Cups
    {
        Dictionary<string, EndesaEntity.Cups> dic_cups = 
            new Dictionary<string, EndesaEntity.Cups>();

        public CupsFunciones()
        {
            CargaInventario();
        }

        public void GetFromCups20(string cups20)
        {
            EndesaEntity.Cups c;
            if (dic_cups.TryGetValue(cups20, out c))
            {
                this.cups20 = cups20;
                this.cups13 = c.cups13;
                this.id = c.id;
            }
            else
            {
                this.cups20 = cups20;
                this.cups13 = cups20;
                this.id = 0;
            }

        }


        private void CargaInventario()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                strSql = "select c.* from med.adif_cups c";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.Cups p = new EndesaEntity.Cups();
                    p.id = Convert.ToInt32(r["ID_CUPS"]);
                    p.cups13 = r["CUPS13"].ToString();
                    p.cups20 = r["CUPS20"].ToString();

                    EndesaEntity.Cups pp;
                    if (!dic_cups.TryGetValue(p.cups20, out pp))
                        dic_cups.Add(p.cups20, p);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "P01011_Funciones.CargaInventario",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }
    }
}
