using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.eer
{
    public class PreciosEnergia
    {
        Dictionary<string, List<EndesaEntity.PreciosEnergia>> dic;
        public PreciosEnergia(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }
        
        private Dictionary<string, List<EndesaEntity.PreciosEnergia>> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, List<EndesaEntity.PreciosEnergia>> d
                = new Dictionary<string, List<EndesaEntity.PreciosEnergia>>();

            try
            {
                strSql = "SELECT cups20, fecha_desde, fecha_hasta,"
                    + " p1, p2, p3, p4, p5, p6, descuentos_te"
                    + " FROM cont.eer_precios_energia";

                if(fd != DateTime.MinValue && fh != DateTime.MaxValue)
                {
                    strSql += " where (fecha_desde <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_hasta >= '" + fd.ToString("yyyy-MM-dd") + "')";
                }
                
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.PreciosEnergia c = new EndesaEntity.PreciosEnergia();
                    c.cups20 = r["cups20"].ToString();
                    c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);
                    c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);
                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["p" + i] != System.DBNull.Value)
                            c.precios_periodo[i] = Convert.ToDouble(r["p" + i]);
                    }
                    if (r["descuentos_te"] != System.DBNull.Value)
                        c.descuento_te = Convert.ToDouble(r["descuentos_te"]);

                    List<EndesaEntity.PreciosEnergia> o;
                    if (!d.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.PreciosEnergia>();
                        o.Add(c);
                        d.Add(c.cups20, o);
                    }
                    else
                        o.Add(c);
                        
                }
                db.CloseConnection();
                return d;
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }


        //public EndesaEntity.PreciosEnergia GetPrecio(string cups20)
        //{
        //    List<EndesaEntity.PreciosEnergia> o;
        //    if (dic.TryGetValue(cups20, out o))
        //        return o;
        //    else
        //        return null;
        //}

        public EndesaEntity.PreciosEnergia GetPrecio(string cups20, DateTime fd, DateTime fh)
        {
            List<EndesaEntity.PreciosEnergia> o;
            EndesaEntity.PreciosEnergia p;
            if (dic.TryGetValue(cups20, out o))
            {
                for(int i = 0; i < o.Count(); i++)
                {
                    if(o[i].fecha_desde <= fd && o[i].fecha_hasta >= fh)
                        return o[i];
                }
                return null;
            }                
            else
                return null;
        }
    }
}
