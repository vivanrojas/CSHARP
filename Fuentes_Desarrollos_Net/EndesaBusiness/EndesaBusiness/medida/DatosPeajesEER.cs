using EndesaBusiness.servidores;
using EndesaEntity.eer;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class DatosPeajesEER
    {

        Dictionary<string, EndesaEntity.DatosPeaje> dic;
        public DatosPeajesEER(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }

        private Dictionary<string, EndesaEntity.DatosPeaje> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.DatosPeaje> d
                = new Dictionary<string, EndesaEntity.DatosPeaje>();

            try
            {
                strSql = "SELECT cups20, fd, fh, version, cod_fiscal, fecha_factura,"
                    + " a_p1, a_p2, a_p3, a_p4, a_p5, a_p6, r_p1, r_p2, r_p3, r_p4, r_p5, r_p6,"
                    + " potmax_1, potmax_2, potmax_3, potmax_4, potmax_5, potmax_6,"
                    + " importe_termino_potencia, importe_excesos_potencia, importe_excesos_reactiva, f_ult_mod"
                    + " FROM cont.eer_datos_medida"
                    + " where (fd <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fh >= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.DatosPeaje c = new EndesaEntity.DatosPeaje();
                    c.cups20 = r["cups20"].ToString();
                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["a_p" + i] != System.DBNull.Value)
                        {
                            c.activa[i] = Convert.ToDouble(r["a_p" + i]);
                            c.total_energia_activa = c.total_energia_activa + c.activa[i];
                        }
                             

                        if (r["r_p" + i] != System.DBNull.Value)
                        {
                            c.reactiva[i] = Convert.ToDouble(r["r_p" + i]);
                            c.total_energia_reactiva = c.total_energia_reactiva + c.reactiva[i];
                        }
                            

                        if (r["potmax_" + i] != System.DBNull.Value)
                            c.reactiva[i] = Convert.ToDouble(r["potmax_" + i]);

                    }

                    if (r["importe_termino_potencia"] != System.DBNull.Value)
                        c.importe_termino_potencia = Convert.ToDouble(r["importe_termino_potencia"]);

                    if (r["importe_excesos_potencia"] != System.DBNull.Value)
                        c.importe_excesos_potencia = Convert.ToDouble(r["importe_excesos_potencia"]);

                    if (r["importe_excesos_reactiva"] != System.DBNull.Value)
                        c.importe_excesos_reactiva = Convert.ToDouble(r["importe_excesos_reactiva"]);

                    d.Add(c.cups20, c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public EndesaEntity.DatosPeaje GetDatosPeaje(string cups20)
        {
            EndesaEntity.DatosPeaje o;
            if(dic.TryGetValue(cups20, out o))
            {
                return o;
            }else
                return null;

        }

        public bool HayDatos(string cups20)
        {
            EndesaEntity.DatosPeaje o;
            if(dic != null)
            {
                return (dic.TryGetValue(cups20, out o));
            }else
                return false;
            
        }

    }
}
