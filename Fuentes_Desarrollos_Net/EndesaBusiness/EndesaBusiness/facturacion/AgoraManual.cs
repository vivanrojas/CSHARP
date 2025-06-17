using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class AgoraManual
    {
        Dictionary<string, EndesaEntity.facturacion.AgoraManual> dic;

        public AgoraManual()
        {
            dic = Carga(DateTime.MinValue, DateTime.MinValue);
        }

        public AgoraManual(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }

        private Dictionary<string, EndesaEntity.facturacion.AgoraManual> Carga(DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, EndesaEntity.facturacion.AgoraManual> d =
                new Dictionary<string, EndesaEntity.facturacion.AgoraManual>();


            strSql = "SELECT Empresa, NIF, CCOUNIPS, CUPS20, DAPERSOC, GRUPO, FD, FH, PRECIOS, FACTURAS_A_CUENTA"
                + " FROM fact.cm_sofisticados";

            if(fd != DateTime.MinValue)
            {
                strSql += " WHERE (FD <= '" + fd.ToString("yyyy-MM-dd") + "' AND"
                    + " FH >= '" + fh.ToString("yyyy-MM-dd") + "')";
            }

            strSql += " order by CUPS20, FD DESC";   
            

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                EndesaEntity.facturacion.AgoraManual c = new EndesaEntity.facturacion.AgoraManual();

                if (r["Empresa"] != System.DBNull.Value)
                    c.empresa = r["Empresa"].ToString();

                if (r["NIF"] != System.DBNull.Value)
                    c.nif = r["NIF"].ToString();

                if (r["CCOUNIPS"] != System.DBNull.Value)
                    c.cups13 = r["CCOUNIPS"].ToString();

                if (r["CUPS20"] != System.DBNull.Value)
                    c.cups20 = r["CUPS20"].ToString();

                if (r["DAPERSOC"] != System.DBNull.Value)
                    c.cliente = r["DAPERSOC"].ToString();

                if (r["GRUPO"] != System.DBNull.Value)
                    c.grupo = r["GRUPO"].ToString();

                if (r["FD"] != System.DBNull.Value)
                    c.fecha_vigor_desde = Convert.ToDateTime(r["FD"]);

                if (r["FH"] != System.DBNull.Value)
                    c.fecha_vigor_hasta = Convert.ToDateTime(r["FH"]);

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
