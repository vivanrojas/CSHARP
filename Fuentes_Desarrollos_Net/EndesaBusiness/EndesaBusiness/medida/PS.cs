using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class PS
    {
        public Dictionary<string, EndesaEntity.medida.Ps_mstr> dic = new Dictionary<string, EndesaEntity.medida.Ps_mstr>();

        public PS()
        {
            dic = new Dictionary<string, EndesaEntity.medida.Ps_mstr>();

        }
        public void Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "select IDMP, CUPSREE from ps_mstr";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Ps_mstr c = new EndesaEntity.medida.Ps_mstr();
                    c.idmp = Convert.ToInt32(r["IDMP"]);
                    c.cups22 = r["CUPSREE"].ToString();
                    EndesaEntity.medida.Ps_mstr o;
                    if (!dic.TryGetValue(c.cups22, out o))
                        dic.Add(c.cups22, c);
                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public int Get_IDMP(string cups22)
        {

            EndesaEntity.medida.Ps_mstr o;
            if (dic.TryGetValue(cups22, out o))
                return o.idmp;
            else
                return 0;
        }

        public void Update_ps_mstr()
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql;

            strSql = "insert ignore into ps_mstr (CUPSREE, CCOUNIPS, CUPS20)"
                + " select r.CUPS22, null, substr(r.CUPS22,1,20) as CUPS20 from cchodlast r"
                + " left outer join ps_mstr c on"
                + " c.CUPSREE = r.CUPS22"
                + " where c.CUPSREE is null"
                + " group by r.CUPS22";
            db = new MySQLDB(MySQLDB.Esquemas.MED);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
    }
}
