using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
 
    // Para tratar los CUPS que reporta D&P como
    // AGORA
    public class Agora_DyP
    {

        Dictionary<string, EndesaEntity.facturacion.AgoraManual> dic;

        public Agora_DyP()
        {
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.facturacion.AgoraManual> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, EndesaEntity.facturacion.AgoraManual> d =
                new Dictionary<string, EndesaEntity.facturacion.AgoraManual>();


            strSql = "SELECT CUPS20"
                + " FROM cont.cm_sofisticados_tmp";


            db = new MySQLDB(MySQLDB.Esquemas.CON);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.AgoraManual c = new EndesaEntity.facturacion.AgoraManual();
                

                if (r["CUPS20"] != System.DBNull.Value)
                    c.cups20 = r["CUPS20"].ToString();
                

                EndesaEntity.facturacion.AgoraManual o;
                if (!d.TryGetValue(c.cups20, out o))
                    d.Add(c.cups20, c);

            }
            db.CloseConnection();
            return d;

        }

        public bool EsAgoraManual(string cups20)
        {
            EndesaEntity.facturacion.AgoraManual o;
            return dic.TryGetValue(cups20, out o);
        }
    }
}
