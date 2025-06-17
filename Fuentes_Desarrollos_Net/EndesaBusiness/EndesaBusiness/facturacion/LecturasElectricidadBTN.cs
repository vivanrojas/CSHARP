using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class LecturasElectricidadBTN
    {
        Dictionary<string, List<EndesaEntity.medida.LecturasElectricidad>> dic;
        public LecturasElectricidadBTN()
        {
            dic = new Dictionary<string, List<EndesaEntity.medida.LecturasElectricidad>>();
        }

        public LecturasElectricidadBTN(List<string> lista_cups13, DateTime fd, DateTime fh)
        {
            dic = Carga(lista_cups13, fd, fh);
        }


        private Dictionary<string, List<EndesaEntity.medida.LecturasElectricidad>> Carga(List<string> lista_cups13, DateTime fd, DateTime fh)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;

            Dictionary<string, List<EndesaEntity.medida.LecturasElectricidad>> d =
                new Dictionary<string, List<EndesaEntity.medida.LecturasElectricidad>>();


            try
            {
                strSql = "select l.CUPS, l.FLECTURA, l.TFUENTE, l.Campo16, p.descripcion, p.fuente"
                    + " FROM med_fuentes_lecturas_sce_btn l"
                    + " INNER JOIN med_fuentes_lecturas_sce_btn_p_tfuentes p ON"
                    + " p.id_fuente = l.TFUENTE"
                    + " where (FLECTURA >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FLECTURA <= '" + fh.ToString("yyyy-MM-dd") + "') and"
                    + " CUPS in (";

                for (int i = 0; i < lista_cups13.Count; i++)
                {
                    if (firstOnly)
                    {
                        strSql += "'" + lista_cups13[i] + "'";
                        firstOnly = false;
                    }
                    else                    
                        strSql += " ,'" + lista_cups13[i] + "'";
                    
                }

                strSql += ")"
                    + "ORDER BY l.FLECTURA DESC, l.Campo16 DESC;";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.LecturasElectricidad c = new EndesaEntity.medida.LecturasElectricidad();
                    c.cups13 = r["CUPS"].ToString();
                    c.fecha_lectura = Convert.ToDateTime(r["FLECTURA"]);
                    c.tfuente = Convert.ToInt32(r["TFUENTE"]);
                    c.descripcion_fuente = r["descripcion"].ToString();
                    c.fuente = r["fuente"].ToString();

                    List<EndesaEntity.medida.LecturasElectricidad> o;
                    if (!d.TryGetValue(c.cups13, out o))
                    {
                        o = new List<EndesaEntity.medida.LecturasElectricidad>();
                        o.Add(c);
                        d.Add(c.cups13, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();

                return d;

            }catch(Exception ex)
            {
                return null;
            }
        }


        public EndesaEntity.medida.LecturasElectricidad GetLectura(string ccounips, DateTime flectura)
        {
            List<EndesaEntity.medida.LecturasElectricidad> o;
            if (dic.TryGetValue(ccounips, out o))
            {
                foreach(EndesaEntity.medida.LecturasElectricidad p in o)
                {
                    if (p.fecha_lectura == flectura)
                        return p;
                }
            }
            return null;
        }

    }
}
